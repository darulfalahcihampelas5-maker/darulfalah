import { initializeApp, getApps, getApp } from 'firebase/app';
import { getAuth, GoogleAuthProvider } from 'firebase/auth';
import { 
  initializeFirestore, 
  getFirestore, 
  persistentLocalCache, 
  persistentMultipleTabManager,
  memoryLocalCache,
  Firestore
} from 'firebase/firestore';
import * as firebaseConfigRaw from '../firebase-applet-config.json';

/* eslint-disable-next-line @typescript-eslint/no-explicit-any */
const firebaseConfig = (firebaseConfigRaw as any).default || firebaseConfigRaw;

const app = !getApps().length ? initializeApp(firebaseConfig) : getApp();

let firestoreDb: Firestore;
try {
  firestoreDb = initializeFirestore(app, {
    experimentalAutoDetectLongPolling: true,
    localCache: persistentLocalCache({
      tabManager: persistentMultipleTabManager()
    })
  }, firebaseConfig.firestoreDatabaseId);
} catch {
  try {
    firestoreDb = getFirestore(app, firebaseConfig.firestoreDatabaseId);
  } catch {
    try {
      firestoreDb = initializeFirestore(app, {
        experimentalAutoDetectLongPolling: true,
        localCache: memoryLocalCache()
      }, firebaseConfig.firestoreDatabaseId);
    } catch {
      firestoreDb = getFirestore(app);
    }
  }
}

export const db = firestoreDb;
export const dbDefault = firestoreDb;
export const auth = getAuth(app);
export const googleProvider = new GoogleAuthProvider();


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
  }
}

export function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errorMessage = error instanceof Error ? error.message : String(error);
  
  // Ignore harmless transient errors like database closing/hidden or background sync notices
  if (
    errorMessage.toLowerCase().includes('closing') ||
    errorMessage.toLowerCase().includes('hidden') ||
    errorMessage.toLowerCase().includes('invalidstateerror') ||
    errorMessage.toLowerCase().includes('sync background') ||
    errorMessage.toLowerCase().includes('background sync') ||
    errorMessage.toLowerCase().includes('unavailable') ||
    errorMessage.toLowerCase().includes('could not reach cloud firestore') ||
    errorMessage.toLowerCase().includes('offline mode') ||
    errorMessage.toLowerCase().includes('client is offline')
  ) {
    return;
  }

  const errInfo: FirestoreErrorInfo = {
    error: errorMessage,
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

  // Detect quota exhausted
  if (errorMessage.toLowerCase().includes('quota') || errorMessage.toLowerCase().includes('resource-exhausted')) {
    window.dispatchEvent(new CustomEvent('firestore-quota-exceeded', { detail: errInfo }));
  }
  
  if (errorMessage.toLowerCase().includes('permission') || errorMessage.toLowerCase().includes('missing or insufficient')) {
    throw new Error(JSON.stringify(errInfo));
  }
}

export async function testConnection() {
  try {
    // Optional connection check
  } catch {
    // Ignore
  }
}

