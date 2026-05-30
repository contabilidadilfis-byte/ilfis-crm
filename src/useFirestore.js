import { useState, useEffect, useCallback } from 'react';
import { collection, onSnapshot, addDoc, updateDoc, deleteDoc, doc, serverTimestamp } from 'firebase/firestore';
import { db } from './firebase';

const now = new Date();
const daysAgo = (d) => new Date(now - d * 86400000).toISOString().split('T')[0];

const SEED_ADVISORS = [
  { name: 'Laura Gómez', email: 'l.gomez@ilfis.com', phone: '+56 9 1111 2222', role: 'Senior', active: true },
  { name: 'Carlos Mendoza', email: 'c.mendoza@ilfis.com', phone: '+52 55 3333 4444', role: 'Senior', active: true },
  { name: 'Ana Ríos', email: 'a.rios@ilfis.com', phone: '+51 1 5555 6666', role: 'Junior', active: true },
  { name: 'Pedro Salinas', email: 'p.salinas@ilfis.com', phone: '+57 1 7777 8888', role: 'Junior', active: true },
  { name: 'María Torres', email: 'm.torres@ilfis.com', phone: '+54 11 9999 0000', role: 'Senior', active: true },
];

const SEED_CLIENTS = [
  { company: 'TechNova S.A.', contact: 'Rodrigo Fuentes', phone: '+56 9 1234 5678', email: 'r.fuentes@technova.cl', country: 'Chile', product: 'CRM Enterprise', value: 8500, advisor: 'Laura Gómez', stage: 'En Negociación', createdAt: daysAgo(75) },
  { company: 'DataSoft Perú', contact: 'Luis Vargas', phone: '+51 1 9876 5432', email: 'lvargas@datasoft.pe', country: 'Perú', product: 'Analytics Pro', value: 12000, advisor: 'Ana Ríos', stage: 'Ganado', createdAt: daysAgo(40) },
  { company: 'Consultora BCD', contact: 'Andrea Luna', phone: '+56 2 3456 7890', email: 'aluna@bcd.cl', country: 'Chile', product: 'ERP Avanzado', value: 18000, advisor: 'Laura Gómez', stage: 'Ganado', createdAt: daysAgo(60) },
];

export function useFirestore() {
  const [clients, setClients] = useState([]);
  const [payments, setPayments] = useState([]);
  const [advisors, setAdvisors] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);
  const [seeded, setSeeded] = useState(false);

  const seedIfEmpty = useCallback(async (advisorDocs, clientDocs) => {
    if (seeded) return;
    setSeeded(true);
    try {
      if (false && advisorDocs.length === 0) {
        for (const a of SEED_ADVISORS) {
          await addDoc(collection(db, 'advisors'), { ...a, createdAt: serverTimestamp() });
        }
      }
      if (false && clientDocs.length === 0) {
        for (let i = 0; i < SEED_CLIENTS.length; i++) {
          await addDoc(collection(db, 'clients'), { ...SEED_CLIENTS[i], _seededIdx: i + 1 });
        }
      }
    } catch (err) {
      console.warn('Seed error:', err.message);
    }
  }, [seeded]);

  useEffect(() => {
    let advisorUnsub, clientUnsub, paymentUnsub;
    let advisorData = [], clientData = [];
    const handleError = (err) => { setError(err.message); setLoading(false); };

    advisorUnsub = onSnapshot(collection(db, 'advisors'), (snap) => {
      advisorData = snap.docs.map(d => ({ ...d.data(), id: d.id }));
      setAdvisors(advisorData);
    }, handleError);

    clientUnsub = onSnapshot(collection(db, 'clients'), (snap) => {
      clientData = snap.docs.map(d => ({ ...d.data(), id: d.id }));
      setClients(clientData);
      if (!seeded) seedIfEmpty(advisorData, clientData);
      setLoading(false);
    }, handleError);

    paymentUnsub = onSnapshot(collection(db, 'payments'), (snap) => {
      setPayments(snap.docs.map(d => ({ ...d.data(), id: d.id })));
    }, handleError);

    return () => { advisorUnsub && advisorUnsub(); clientUnsub && clientUnsub(); paymentUnsub && paymentUnsub(); };
  }, []);

  const saveClient = useCallback(async (data, existingId = null) => {
    const clean = { ...data }; delete clean.id;
    if (existingId) await updateDoc(doc(db, 'clients', existingId), { ...clean, updatedAt: serverTimestamp() });
    else await addDoc(collection(db, 'clients'), { ...clean, createdAt: new Date().toISOString().split('T')[0] });
  }, []);

  const deleteClient = useCallback(async (id) => { await deleteDoc(doc(db, 'clients', id)); }, []);

  const savePayment = useCallback(async (data, existingId = null) => {
    const clean = { ...data }; delete clean.id;
    if (existingId) await updateDoc(doc(db, 'payments', existingId), { ...clean, updatedAt: serverTimestamp() });
    else await addDoc(collection(db, 'payments'), { ...clean, createdAt: serverTimestamp() });
  }, []);

  const saveAdvisor = useCallback(async (data, existingId = null) => {
    const clean = { ...data }; delete clean.id;
    if (existingId) await updateDoc(doc(db, 'advisors', existingId), clean);
    else await addDoc(collection(db, 'advisors'), { ...clean, createdAt: serverTimestamp() });
  }, []);

  const deleteAdvisor = useCallback(async (id) => { await deleteDoc(doc(db, 'advisors', id)); }, []);

const lockClient = useCallback(async (clientId, advisorName) => {
  await updateDoc(doc(db, 'clients', clientId), {
    lockedBy: advisorName,
    lockedAt: serverTimestamp(),
  });
}, []);

const unlockClient = useCallback(async (clientId) => {
  await updateDoc(doc(db, 'clients', clientId), {
    lockedBy: null,
    lockedAt: null,
  });
}, []);

const isClientLocked = useCallback((client, currentAdvisor) => {
  if (!client.lockedBy) return false;
  if (client.lockedBy === currentAdvisor) return false;
  if (client.lockedAt?.toDate) {
    const diff = Date.now() - client.lockedAt.toDate().getTime();
    if (diff > 30 * 60 * 1000) return false;
  }
  return true;
}, []);

return { clients, payments, advisors, loading, error, saveClient, deleteClient, savePayment, saveAdvisor, deleteAdvisor, lockClient, unlockClient, isClientLocked };
