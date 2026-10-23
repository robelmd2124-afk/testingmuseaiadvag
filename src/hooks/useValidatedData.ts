import { useState, useEffect, useRef, useCallback } from 'react';
import { SyncLogger } from '../services/syncLogger';

/**
 * Data synchronization lifecycle states
 */
export type SyncState = 'IDLE' | 'LOADING' | 'VALID_DATA' | 'EMPTY_DATA' | 'ERROR' | 'REVERTED_TO_LAST_KNOWN';

/**
 * Configuration options for the useValidatedData hook
 */
export interface UseValidatedDataOptions<T> {
  /** Initial fallback data if no previous valid data exists */
  initialData?: T;
  /**
   * Custom validator function to determine whether the incoming data is valid.
   * If not provided, a strict default validator checks for non-null, non-empty, and positive numeric values.
   */
  validator?: (data: T, previousValidData: T | null) => boolean;
  /** Unique key for caching / persistence across renders */
  key?: string;
  /** React dependency array to trigger refetch when values change */
  dependencies?: any[];
  /** Whether the initial fetch is enabled (default: true) */
  enabled?: boolean;
  /** Callback invoked when a sync issue occurs and data reverts to last known state */
  onSyncIssue?: (warningMessage: string, rawData: any) => void;
  /** Callback invoked when new valid data is confirmed and committed */
  onValidData?: (validData: T) => void;
  /** Auto-clear warning banner timeout in ms (default: 6000ms, 0 to keep active) */
  warningTimeoutMs?: number;
}

/**
 * Return type of useValidatedData hook
 */
export interface UseValidatedDataResult<T> {
  /** The current valid data (guaranteed never to be overwritten by temporary 0 or empty structures) */
  data: T | null;
  /** Direct reference to the last known valid state */
  previousValidData: T | null;
  /** True during initial data loading or when no valid data is available yet */
  isLoading: boolean;
  /** True during background revalidation or manual refresh while preserving current data */
  isSyncing: boolean;
  /** Current synchronization state */
  syncState: SyncState;
  /** Error object if the latest fetch failed */
  error: Error | null;
  /** Warning message if reverted to last known valid state (e.g. 'Sync issue: Reverting to last known state') */
  syncWarning: string | null;
  /** Manually trigger a fresh fetch */
  refetch: () => Promise<T | null>;
  /** Safely mutate or update state with validator enforcement */
  setData: (updater: T | ((prev: T | null) => T)) => boolean;
  /** Dismiss the sync warning banner */
  clearWarning: () => void;
}

/**
 * Default robust validation logic for any data structure.
 * Rejects 0 numbers, empty arrays [], empty objects {}, null, undefined, or invalid payloads.
 */
export function defaultDataValidator<T>(data: T, previousValidData: T | null): boolean {
  if (data === null || data === undefined) {
    return false;
  }

  // Number validation: A raw 0 or NaN is invalid if previous non-zero data existed
  if (typeof data === 'number') {
    if (isNaN(data)) return false;
    if (data === 0 && typeof previousValidData === 'number' && previousValidData > 0) {
      return false; // Prevent temporary zero overwriting real metric
    }
    return true;
  }

  // String validation: Empty string is invalid if previous data existed
  if (typeof data === 'string') {
    if (data.trim() === '' && previousValidData && typeof previousValidData === 'string' && previousValidData.trim() !== '') {
      return false;
    }
    return true;
  }

  // Array validation: An empty array [] is rejected if we previously had valid rows
  if (Array.isArray(data)) {
    if (data.length === 0 && Array.isArray(previousValidData) && previousValidData.length > 0) {
      return false; // Temporary empty array response during sync
    }
    return true;
  }

  // Object validation: Check for empty object {} or objects where all key metrics are 0
  if (typeof data === 'object') {
    const keys = Object.keys(data as object);
    if (keys.length === 0 && previousValidData && typeof previousValidData === 'object' && Object.keys(previousValidData as object).length > 0) {
      return false;
    }

    // Special check for campaign / stats response objects (e.g. { totalRecords: 0, totalBudget: 0 })
    const dObj = data as Record<string, any>;
    const prevObj = (previousValidData && typeof previousValidData === 'object') ? (previousValidData as Record<string, any>) : null;

    if (prevObj && dObj) {
      const hasPrevMetrics = (Number(prevObj.totalRecords || 0) > 0 || Number(prevObj.totalBudget || 0) > 0 || Number(prevObj.totalCampaign || 0) > 0);
      const hasCurrentMetrics = (Number(dObj.totalRecords || 0) > 0 || Number(dObj.totalBudget || 0) > 0 || Number(dObj.totalCampaign || 0) > 0);
      
      // If previous had real metrics but incoming has 0 across all fields without explicit confirmation
      if (hasPrevMetrics && !hasCurrentMetrics && (dObj.success === false || dObj.error || !dObj.confirmedEmpty)) {
        return false;
      }

      // Check if records.rows is empty while previous had rows
      if (Array.isArray(dObj.rows) && dObj.rows.length === 0 && Array.isArray(prevObj.rows) && prevObj.rows.length > 0) {
        return false;
      }
    }

    return true;
  }

  return true;
}

const SYNC_ISSUE_MESSAGE = '';

/**
 * Custom React Hook: `useValidatedData`
 *
 * Wraps any asynchronous data loading state and explicitly maintains a
 * 'previous valid state' reference.
 *
 * If a new fetch returns 0, an empty structure (`[]` / `{}`), null, undefined,
 * or encounters a temporary network/spreadsheet delay, the hook retains the
 * previous valid data and displays a subtle 'Sync issue: Reverting to last known state'
 * error indicator, preventing the UI from rendering 0 values.
 */
export function useValidatedData<T>(
  fetcher: () => Promise<T>,
  options: UseValidatedDataOptions<T> = {}
): UseValidatedDataResult<T> {
  const {
    initialData = null,
    validator = defaultDataValidator,
    key,
    dependencies = [],
    enabled = true,
    onSyncIssue,
    onValidData,
    warningTimeoutMs = 6000
  } = options;

  // Primary state
  const [data, setDataState] = useState<T | null>(() => (initialData !== undefined ? initialData : null));
  const [isLoading, setIsLoading] = useState<boolean>(enabled && initialData === null);
  const [isSyncing, setIsSyncing] = useState<boolean>(false);
  const [syncState, setSyncState] = useState<SyncState>(initialData !== null ? 'VALID_DATA' : 'IDLE');
  const [error, setError] = useState<Error | null>(null);
  const [syncWarning, setSyncWarning] = useState<string | null>(null);

  // Persistent reference to the last known valid state
  const previousValidRef = useRef<T | null>(initialData !== undefined ? initialData : null);
  
  // Request sequence tracker to completely prevent race conditions
  const requestSeqRef = useRef<number>(0);
  const warningTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const mountedRef = useRef<boolean>(true);

  useEffect(() => {
    mountedRef.current = true;
    return () => {
      mountedRef.current = false;
      if (warningTimerRef.current) clearTimeout(warningTimerRef.current);
    };
  }, []);

  const triggerWarning = useCallback((warningMsg: string, rawResult: any) => {
    if (!mountedRef.current) return;
    setSyncWarning(null);
    setSyncState('REVERTED_TO_LAST_KNOWN');
    if (onSyncIssue) onSyncIssue(warningMsg, rawResult);

    if (warningTimerRef.current) clearTimeout(warningTimerRef.current);
    if (warningTimeoutMs > 0) {
      warningTimerRef.current = setTimeout(() => {
        if (mountedRef.current) {
          setSyncWarning(null);
        }
      }, warningTimeoutMs);
    }
  }, [onSyncIssue, warningTimeoutMs]);

  const clearWarning = useCallback(() => {
    if (warningTimerRef.current) clearTimeout(warningTimerRef.current);
    setSyncWarning(null);
  }, []);

  /**
   * Main fetch / sync executor with Last Known Valid Data protection
   */
  const executeFetch = useCallback(async (): Promise<T | null> => {
    if (!enabled) return previousValidRef.current;

    const currentSeq = ++requestSeqRef.current;
    const hasExistingData = previousValidRef.current !== null;

    if (hasExistingData) {
      setIsSyncing(true);
    } else {
      setIsLoading(true);
      setSyncState('LOADING');
    }
    setError(null);

    try {
      const rawResult = await fetcher();

      // Discard stale responses from superseded requests (Race condition prevention)
      if (currentSeq !== requestSeqRef.current || !mountedRef.current) {
        return previousValidRef.current;
      }

      const isValid = validator(rawResult, previousValidRef.current);

      if (isValid) {
        // Confirmed valid data
        previousValidRef.current = rawResult;
        setDataState(rawResult);
        setSyncState('VALID_DATA');
        setSyncWarning(null);
        setError(null);
        if (key) {
          SyncLogger.logSuccess('LOCAL_CACHE', `Validated data committed for ${key}`);
        }
        if (onValidData) onValidData(rawResult);
        return rawResult;
      } else {
        // Data failed validation (returned 0, empty structure, or corrupted payload)
        if (previousValidRef.current !== null) {
          // Revert to last known valid state
          setDataState(previousValidRef.current);
          SyncLogger.logReverted('STATE_REVERT', `Sync issue on ${key || 'data'}: Reverting to last known state to prevent 0 values`, { rawResult });
          triggerWarning(SYNC_ISSUE_MESSAGE, rawResult);
          return previousValidRef.current;
        } else {
          // No prior data existed
          setDataState(rawResult);
          setSyncState('EMPTY_DATA');
          SyncLogger.logWarning('STATE_REVERT', `Received empty structure for ${key || 'data'} (no previous state existed)`);
          return rawResult;
        }
      }
    } catch (err: any) {
      if (currentSeq !== requestSeqRef.current || !mountedRef.current) {
        return previousValidRef.current;
      }

      const errorObj = err instanceof Error ? err : new Error(String(err?.message || err || 'Fetch failed'));
      setError(errorObj);

      if (previousValidRef.current !== null) {
        // Network/API error: retain previous valid data
        setDataState(previousValidRef.current);
        SyncLogger.logReverted('STATE_REVERT', `Network error on ${key || 'data'}: Retaining previous valid state`, errorObj);
        triggerWarning(SYNC_ISSUE_MESSAGE, err);
        return previousValidRef.current;
      } else {
        setSyncState('ERROR');
        SyncLogger.logFailure('RPC_READ', errorObj, { key });
        return null;
      }
    } finally {
      if (mountedRef.current && currentSeq === requestSeqRef.current) {
        setIsLoading(false);
        setIsSyncing(false);
      }
    }
  }, [enabled, fetcher, validator, onValidData, triggerWarning]);

  // Safe manual setter with validation guard
  const setData = useCallback((updater: T | ((prev: T | null) => T)): boolean => {
    const nextVal = typeof updater === 'function' ? (updater as any)(previousValidRef.current) : updater;
    const isValid = validator(nextVal, previousValidRef.current);

    if (isValid) {
      previousValidRef.current = nextVal;
      setDataState(nextVal);
      setSyncState('VALID_DATA');
      setSyncWarning(null);
      setError(null);
      return true;
    } else {
      triggerWarning(SYNC_ISSUE_MESSAGE, nextVal);
      return false;
    }
  }, [validator, triggerWarning]);

  // Lifecycle trigger on dependency change or enable
  useEffect(() => {
    if (enabled) {
      executeFetch();
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [enabled, key, ...dependencies]);

  return {
    data,
    previousValidData: previousValidRef.current,
    isLoading,
    isSyncing,
    syncState,
    error,
    syncWarning,
    refetch: executeFetch,
    setData,
    clearWarning
  };
}

export default useValidatedData;
