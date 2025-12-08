import React, { useEffect, useState } from 'react';
import { getFirestore, collection, getDocs, addDoc, updateDoc, deleteDoc, doc, query, orderBy } from 'firebase/firestore';
import { initializeApp, getApps } from 'firebase/app';

// TODO: Replace with your actual Firebase config
const firebaseConfig = {
  apiKey: "YOUR_API_KEY",
  authDomain: "YOUR_AUTH_DOMAIN",
  projectId: "foreclosure-15f09",
  storageBucket: "YOUR_STORAGE_BUCKET",
  messagingSenderId: "YOUR_MESSAGING_SENDER_ID",
  appId: "YOUR_APP_ID"
};

const app = getApps().length ? getApps()[0] : initializeApp(firebaseConfig);
const db = getFirestore(app);

export default function GlobalParameterTable() {
  const [params, setParams] = useState([]);
  const [loading, setLoading] = useState(true);
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({ param: '', paramValue: '' });
  const [sortCol, setSortCol] = useState('param');
  const [sortDir, setSortDir] = useState('asc');

  useEffect(() => {
    fetchParams();
  }, []);

  async function fetchParams() {
    setLoading(true);
  const q = query(collection(db, 'parameters'), orderBy(sortCol, sortDir));
    const querySnapshot = await getDocs(q);
    const data = querySnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    setParams(data);
    setLoading(false);
  }

  function startEdit(param) {
    setEditing(param.id);
    setForm({ param: param.param, paramValue: param.paramValue });
  }

  function cancelEdit() {
    setEditing(null);
    setForm({ param: '', paramValue: '' });
  }

  async function saveEdit() {
  const ref = doc(db, 'parameters', editing);
    await updateDoc(ref, form);
    setEditing(null);
    fetchParams();
  }

  async function handleDelete(id) {
  await deleteDoc(doc(db, 'parameters', id));
    fetchParams();
  }

  async function handleAdd() {
  await addDoc(collection(db, 'parameters'), form);
    setForm({ param: '', paramValue: '' });
    fetchParams();
  }

  function handleChange(e) {
    setForm({ ...form, [e.target.name]: e.target.value });
  }

  function handleSort(col) {
    if (sortCol === col) {
      setSortDir(sortDir === 'asc' ? 'desc' : 'asc');
    } else {
      setSortCol(col);
      setSortDir('asc');
    }
    setTimeout(fetchParams, 0);
  }

  if (loading) return <div>Loading...</div>;

  return (
    <div style={{ padding: 0 }}>
      <h3>Global Parameters</h3>
      <table border="1" cellPadding="6" style={{ minWidth: 500, borderCollapse: 'separate', background: '#fff', color: '#000' }}>
        <thead>
          <tr>
            <th style={{ cursor: 'pointer', color: '#000' }} onClick={() => handleSort('param')}>
              Param{sortCol === 'param' ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
            </th>
            <th style={{ cursor: 'pointer', color: '#000' }} onClick={() => handleSort('paramValue')}>
              Value{sortCol === 'paramValue' ? (sortDir === 'asc' ? ' ▲' : ' ▼') : ''}
            </th>
            <th style={{ color: '#000' }}>Actions</th>
          </tr>
        </thead>
        <tbody>
          {params.map(param => (
            <tr key={param.id}>
              <td>
                {editing === param.id ? (
                  <input name="param" value={form.param} onChange={handleChange} />
                ) : (
                  param.param
                )}
              </td>
              <td>
                {editing === param.id ? (
                  <input name="paramValue" value={form.paramValue} onChange={handleChange} />
                ) : (
                  param.paramValue
                )}
              </td>
              <td>
                {editing === param.id ? (
                  <>
                    <button onClick={saveEdit}>Save</button>
                    <button onClick={cancelEdit}>Cancel</button>
                  </>
                ) : (
                  <>
                    <button onClick={() => startEdit(param)}>Edit</button>
                    <button onClick={() => handleDelete(param.id)}>Delete</button>
                  </>
                )}
              </td>
            </tr>
          ))}
        </tbody>
      </table>
      <h4 style={{ marginTop: 24 }}>Add New Parameter</h4>
      <form onSubmit={e => { e.preventDefault(); handleAdd(); }}>
        <input
          name="param"
          placeholder="Param"
          value={form.param}
          onChange={handleChange}
          style={{ marginRight: 8 }}
        />
        <input
          name="paramValue"
          placeholder="Value"
          value={form.paramValue}
          onChange={handleChange}
          style={{ marginRight: 8 }}
        />
        <button type="submit">Add</button>
      </form>
    </div>
  );
}
