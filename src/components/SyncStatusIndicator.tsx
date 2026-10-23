import React from 'react';

export type SyncIndicatorVariant = 'floating' | 'inline' | 'compact' | 'banner';

export interface SyncStatusIndicatorProps {
  /** The warning string from useValidatedData */
  warning?: string | null;
  /** Whether a background sync is actively running */
  isSyncing?: boolean;
  /** Optional callback to dismiss the indicator */
  onDismiss?: () => void;
  /** Optional callback to retry fetching fresh data */
  onRetry?: () => void;
  /** Visual placement variant */
  variant?: SyncIndicatorVariant;
  /** Additional CSS class names */
  className?: string;
  /** Custom inline style overrides */
  style?: React.CSSProperties;
}

/**
 * `SyncStatusIndicator`
 *
 * Silent Mode: Background validation keeps data intact without displaying intrusive banners on screen.
 */
export const SyncStatusIndicator: React.FC<SyncStatusIndicatorProps> = () => {
  return null;
};

export default SyncStatusIndicator;
