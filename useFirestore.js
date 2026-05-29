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
  { company: 'DataSoft Perú', contact: 'Luis Vargas', phone: '+51 1 9876 5432', email: 'lvargas@datasoft.pe', country: 'Perú', produc
