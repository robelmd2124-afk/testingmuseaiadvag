import { initializeApp } from 'firebase/app';
import {
  getAuth,
  signInWithPopup,
  GoogleAuthProvider,
  onAuthStateChanged,
  signOut,
  type User
} from 'firebase/auth';
import { getFirestore } from 'firebase/firestore';
import firebaseConfig from '../../firebase-applet-config.json';

const app = initializeApp(firebaseConfig);
export const auth = getAuth(app);
export const db = getFirestore(app, (firebaseConfig as any).firestoreDatabaseId); /* CRITICAL: The app will break without this line */


const provider = new GoogleAuthProvider();
// Request Google Workspace Sheets Scope
export const WORKSPACE_SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets'
];

WORKSPACE_SCOPES.forEach(scope => provider.addScope(scope));

let isSigningIn = false;
// In-memory access token cache
let cachedAccessToken: string | null = null;
let currentUser: User | null = null;

export const initGoogleAuth = (
  onSuccess?: (user: User, token: string | null) => void,
  onFailure?: () => void
) => {
  return onAuthStateChanged(auth, async (user: User | null) => {
    currentUser = user;
    if (user) {
      if (cachedAccessToken) {
        if (onSuccess) onSuccess(user, cachedAccessToken);
      } else if (!isSigningIn) {
        cachedAccessToken = null;
        if (onFailure) onFailure();
      }
    } else {
      cachedAccessToken = null;
      currentUser = null;
      if (onFailure) onFailure();
    }
  });
};

export const googleSignIn = async (): Promise<{ user: User; accessToken: string } | null> => {
  // First try Firebase Auth popup
  try {
    isSigningIn = true;
    const result = await signInWithPopup(auth, provider);
    const credential = GoogleAuthProvider.credentialFromResult(result);
    if (!credential?.accessToken) {
      throw new Error('Google Sign-in did not return an OAuth access token.');
    }
    cachedAccessToken = credential.accessToken;
    currentUser = result.user;
    return { user: result.user, accessToken: cachedAccessToken };
  } catch (error: any) {
    console.warn('Firebase signInWithPopup note:', error?.code || error?.message);
    
    // Check if domain is unauthorized in Firebase console
    const isDomainError = error?.code === 'auth/unauthorized-domain' || 
                          error?.message?.includes('unauthorized-domain');

    // Try Google Identity Services (GSI) as robust fallback
    const clientId = (firebaseConfig as any).oAuthClientId;
    if (clientId) {
      try {
        // Wait briefly if GSI is still finishing script loading
        let attempts = 0;
        while (typeof (window as any).google?.accounts?.oauth2 === 'undefined' && attempts < 10) {
          await new Promise(r => setTimeout(r, 150));
          attempts++;
        }

        if (typeof (window as any).google?.accounts?.oauth2 !== 'undefined') {
          return await new Promise<{ user: any; accessToken: string }>((resolve, reject) => {
            const tokenClient = (window as any).google.accounts.oauth2.initTokenClient({
              client_id: clientId,
              scope: WORKSPACE_SCOPES.join(' '),
              callback: async (resp: any) => {
                if (resp.error) {
                  return reject(new Error(resp.error_description || resp.error));
                }
                const token = resp.access_token;
                cachedAccessToken = token;
                let userInfo: any = { email: 'Google User', displayName: 'Google Account' };
                try {
                  const uRes = await fetch('https://www.googleapis.com/oauth2/v3/userinfo', {
                    headers: { Authorization: `Bearer ${token}` }
                  });
                  if (uRes.ok) {
                    const uData = await uRes.json();
                    userInfo = { email: uData.email, displayName: uData.name };
                    try {
                      localStorage.setItem('adverse_google_user_email', uData.email || '');
                      localStorage.setItem('adverse_google_user_name', uData.name || '');
                    } catch (_) {}
                  }
                } catch (e) {}
                currentUser = userInfo as any;
                resolve({ user: userInfo, accessToken: token });
              }
            });
            tokenClient.requestAccessToken({ prompt: 'consent' });
          });
        }
      } catch (gsiErr: any) {
        console.warn('GSI fallback also note:', gsiErr?.message);
      }
    }

    if (isDomainError) {
      const friendlyError: any = new Error('UNAUTHORIZED_DOMAIN');
      friendlyError.code = 'auth/unauthorized-domain';
      throw friendlyError;
    }
    throw error;
  } finally {
    isSigningIn = false;
  }
};

export const getGoogleAccessToken = (): string | null => {
  return cachedAccessToken;
};

export const getCurrentGoogleUser = (): User | null => {
  return currentUser;
};

export const googleSignOut = async (): Promise<void> => {
  await signOut(auth);
  cachedAccessToken = null;
  currentUser = null;
};
