import React, { useEffect, useState, useCallback, useRef } from 'react';
import { useValidatedData } from './hooks/useValidatedData';
import { SyncStatusIndicator } from './components/SyncStatusIndicator';

export { useValidatedData };
export { SyncStatusIndicator };

export interface FocusedRowState {
  index: number | null;
  record: Record<string, any> | null;
  totalRecords: number;
}

/**
 * Custom hook for managing Table Row Focus Mode and the Floating Action Menu.
 * Enables clicking any table row to focus and select the underlying campaign record,
 * allowing floating actions (Edit, Delete, Copy, Increase, WhatsApp, Clone, Inspect)
 * to operate on the correct record.
 */
export function useRowFocusMode() {
  const [focusedRow, setFocusedRow] = useState<FocusedRowState>({
    index: null,
    record: null,
    totalRecords: 0
  });

  const clearFocus = useCallback(() => {
    setFocusedRow({ index: null, record: null, totalRecords: 0 });
    if (typeof (window as any).clearRowFocus === 'function') {
      (window as any).clearRowFocus();
    }
  }, []);

  const selectRow = useCallback((index: number, shouldScroll: boolean = false) => {
    const globalState = (window as any).state;
    const records =
      globalState?.currentRenderedRecords ||
      globalState?.filteredRecords ||
      globalState?.rawRecords ||
      [];

    if (!records.length || index < 0 || index >= records.length) {
      return;
    }

    const record = records[index];
    if (!record) return;

    // Synchronize global application state so Vanilla JS actions (Edit, Delete, Copy) target this record
    if (globalState) {
      globalState.focusedRowIndex = index;
      globalState.editingRow = record;
      globalState.selectedRecord = record;
      globalState.editingId =
        record['Record ID'] ||
        record._recordId ||
        record._row ||
        record['SL NO'] ||
        record.sl_no ||
        record.SL ||
        record.Date;
    }

    setFocusedRow({
      index,
      record,
      totalRecords: records.length
    });

    // Directly open the Campaign Inspector modal for this row
    if (typeof (window as any).openCampaignInspector === 'function') {
      (window as any).openCampaignInspector(index);
    }
  }, []);

  // Action Dispatchers for the focused record
  const executeEdit = useCallback(() => {
    if (typeof (window as any).rfActionEdit === 'function') {
      (window as any).rfActionEdit();
    }
  }, []);

  const executeDelete = useCallback(() => {
    if (typeof (window as any).rfActionDelete === 'function') {
      (window as any).rfActionDelete();
    }
  }, []);

  const executeCopy = useCallback(() => {
    if (typeof (window as any).rfActionCopy === 'function') {
      (window as any).rfActionCopy();
    }
  }, []);

  const executeIncrease = useCallback(() => {
    if (typeof (window as any).rfActionIncrease === 'function') {
      (window as any).rfActionIncrease();
    }
  }, []);

  const executeWhatsApp = useCallback(() => {
    if (typeof (window as any).rfActionWhatsApp === 'function') {
      (window as any).rfActionWhatsApp();
    }
  }, []);

  const executeClone = useCallback(() => {
    if (typeof (window as any).rfActionClone === 'function') {
      (window as any).rfActionClone();
    }
  }, []);

  const executeInspect = useCallback(() => {
    if (typeof (window as any).rfActionInspect === 'function') {
      (window as any).rfActionInspect();
    }
  }, []);

  return {
    focusedIndex: focusedRow.index,
    selectedRecord: focusedRow.record,
    totalRecords: focusedRow.totalRecords,
    isFocused: focusedRow.index !== null,
    selectRow,
    clearFocus,
    executeEdit,
    executeDelete,
    executeCopy,
    executeIncrease,
    executeWhatsApp,
    executeClone,
    executeInspect
  };
}

export default function App() {
  const { isFocused, selectRow, clearFocus } = useRowFocusMode();

  // Validate and preserve settings state across refreshes
  const {
    isSyncing,
    syncWarning,
    clearWarning
  } = useValidatedData(
    async () => {
      try {
        const res = await fetch('/api/rpc?method=getSettings', { credentials: 'include' });
        const json = await res.json();
        return json?.result || null;
      } catch (e) {
        return null;
      }
    },
    {
      key: 'app_settings',
      validator: (data) => Boolean(data && typeof data === 'object')
    }
  );

  // Global Table Click Listener: Intercepts table row clicks to activate Row Focus Mode & Floating Menu
  useEffect(() => {
    const handleGlobalTableClick = (e: MouseEvent) => {
      const target = e.target as HTMLElement | null;
      if (!target) return;

      // Ignore interactive controls inside row (buttons, actions, inputs, links, dropdowns)
      if (
        target.closest('.td-action') ||
        target.closest('button') ||
        target.closest('input') ||
        target.closest('select') ||
        target.closest('a') ||
        target.closest('.page-hover-popover') ||
        target.closest('#rowFocusFloatingMenu') ||
        target.closest('.modal-overlay')
      ) {
        return;
      }

      // Check if clicked element is inside a table record row
      const tr = target.closest('tr.table-record-row, tr[data-record-idx]') as HTMLTableRowElement | null;
      if (tr) {
        // Only Admins can activate row focus and campaign inspector; viewers only see read-only data
        const isAdmin = typeof (window as any).isAdmin === 'function' ? (window as any).isAdmin() : false;
        if (!isAdmin) {
          return;
        }

        // Extract row index
        const idxAttr = tr.getAttribute('data-record-idx') || tr.getAttribute('data-idx') || tr.id.replace('tableRecordRow-', '');
        const rowIndex = parseInt(idxAttr, 10);

        if (!isNaN(rowIndex)) {
          e.preventDefault();
          selectRow(rowIndex, false);
        }
      }
    };

    document.addEventListener('click', handleGlobalTableClick);
    return () => {
      document.removeEventListener('click', handleGlobalTableClick);
    };
  }, [selectRow]);

  return (
    <div className="react-bridge-root" style={{ position: 'relative' }}>
      {/* Silent mode: data protection active with zero intrusive warning banners */}
    </div>
  );
}
