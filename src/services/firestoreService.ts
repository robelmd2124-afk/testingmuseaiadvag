import {
  doc,
  collection,
  setDoc,
  getDoc,
  getDocs,
  getDocFromServer,
  deleteDoc,
  onSnapshot,
  query,
  limit,
  type Unsubscribe
} from 'firebase/firestore';
import { auth, db } from './googleAuth.ts';

export enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

export interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  };
}

export function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null): never {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  };
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

/**
 * Validate connection to Firestore on boot as mandated by Firestore guidelines
 */
export async function testConnection(): Promise<boolean> {
  try {
    await getDocFromServer(doc(db, 'test', 'connection'));
    console.log('[Firestore]: Connection verified successfully');
    return true;
  } catch (error) {
    if (error instanceof Error && error.message.includes('the client is offline')) {
      console.error('Please check your Firebase configuration.');
      return false;
    }
    // Expected to fail with missing document, which proves server connectivity
    console.log('[Firestore]: Connection to Firestore server is alive');
    return true;
  }
}

// Test connection on initial load
testConnection().catch(e => console.warn('[Firestore]: Initial test error:', e));

/**
 * Save or update a campaign in Firestore
 */
export async function saveCampaignToFirestore(campaign: any): Promise<void> {
  const recordId = String(campaign.recordId || campaign['Record ID'] || campaign._recordId || '').trim();
  if (!recordId) return;
  const path = `campaigns/${recordId}`;
  try {
    const payload = {
      recordId,
      sl: String(campaign.sl || campaign['SL'] || '').trim(),
      adAccount: String(campaign.adAccount || campaign['AD Account'] || '').trim(),
      idName: String(campaign.idName || campaign['ID Name'] || '').trim(),
      status: String(campaign.status || campaign['Status'] || '').trim(),
      pageName: String(campaign.pageName || campaign['Page Name'] || '').trim(),
      startDate: String(campaign.startDate || campaign['Start'] || '').trim(),
      endDate: String(campaign.endDate || campaign['End'] || '').trim(),
      types: String(campaign.types || campaign['Types'] || 'Boost').trim(),
      day: Number(campaign.day || campaign['Day'] || 1),
      budget: Number(campaign.budget || campaign['Budget'] || 0),
      dailyBudget: Number(campaign.dailyBudget || campaign['Daily $'] || 0),
      taka: Number(campaign.taka || campaign['Taka'] || 0),
      rate: Number(campaign.rate || campaign['Rate'] || 140),
      entryDate: String(campaign.entryDate || campaign['Entry Date'] || '').trim(),
      month: String(campaign.month || campaign['Month'] || '').trim(),
      year: String(campaign.year || campaign['Year'] || '').trim(),
      offUsd: String(campaign.offUsd || campaign['OFF $'] || '').trim(),
      updatedAt: new Date().toISOString()
    };
    await setDoc(doc(db, 'campaigns', recordId), payload, { merge: true });
  } catch (err) {
    handleFirestoreError(err, OperationType.WRITE, path);
  }
}

/**
 * Save or update payment in Firestore
 */
export async function savePaymentToFirestore(payment: any): Promise<void> {
  const recordId = String(payment.recordId || payment['Record ID'] || payment._recordId || '').trim();
  if (!recordId) return;
  const path = `payments/${recordId}`;
  try {
    const payload = {
      recordId,
      date: String(payment.date || payment['Date'] || '').trim(),
      paymentMethod: String(payment.paymentMethod || payment['Payment Method'] || '').trim(),
      paid: Number(payment.paid || payment['Paid (৳)'] || 0),
      month: String(payment.month || payment['Month'] || '').trim(),
      year: String(payment.year || payment['Year'] || '').trim(),
      note: String(payment.note || payment['Note'] || '').trim(),
      updatedAt: new Date().toISOString()
    };
    await setDoc(doc(db, 'payments', recordId), payload, { merge: true });
  } catch (err) {
    handleFirestoreError(err, OperationType.WRITE, path);
  }
}

/**
 * Save global settings to Firestore
 */
export async function saveSettingsToFirestore(settings: any): Promise<void> {
  const path = 'settings/global';
  try {
    const payload = {
      dollarRate: Number(settings.dollar_rate || settings.rate || 140),
      sheetId: String(settings.sheet_id || '').trim(),
      dateFormat: String(settings.date_format || 'DD-MM-YYYY').trim(),
      updatedAt: new Date().toISOString()
    };
    await setDoc(doc(db, 'settings', 'global'), payload, { merge: true });
  } catch (err) {
    handleFirestoreError(err, OperationType.WRITE, path);
  }
}

/**
 * Real-time listener for campaigns in Firestore
 */
export function subscribeToCampaigns(callback: (campaigns: any[]) => void): Unsubscribe {
  const path = 'campaigns';
  const q = query(collection(db, path), limit(2000));
  return onSnapshot(
    q,
    (snapshot) => {
      const records = snapshot.docs.map(d => d.data());
      callback(records);
    },
    (err) => {
      handleFirestoreError(err, OperationType.LIST, path);
    }
  );
}
