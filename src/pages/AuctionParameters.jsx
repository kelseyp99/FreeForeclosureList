// Desired column order (updated)
const COLUMN_ORDER = [
  "County",
  "Days FC",
  "Days TD",
  "Foreclosure",
  "List",
  "Court Docs",
  "PA",
  "PA template",
  "PA template Address",
  "Sale Template",
  "Tax Deed",
  "TaxDeedList",
  "UiPath",
  "UiPathTD",
  "WIXparamID",
  "download file",
  "public records",
  "updateWIX"
];
import React, { useEffect, useState } from 'react';
import { getFirestore, collection, getDocs, addDoc, updateDoc, deleteDoc, doc } from 'firebase/firestore';
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

export default function AuctionParametersPage({ tableMinWidth = 700 }) {
  const [params, setParams] = useState([]);
  const [loading, setLoading] = useState(true);
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({});

  useEffect(() => {
    fetchParams();
  }, []);

  async function fetchParams() {
    setLoading(true);
    const querySnapshot = await getDocs(collection(db, 'auction_parameters'));
    const data = querySnapshot.docs.map(doc => ({ id: doc.id, ...doc.data() }));
    setParams(data);
    setLoading(false);
  }

  function startEdit(param) {
    setEditing(param.id);
    setForm(param);
  }

  function cancelEdit() {
    setEditing(null);
    setForm({});
  }

  async function saveEdit() {
    const ref = doc(db, 'auction_parameters', editing);
    await updateDoc(ref, form);
    setEditing(null);
    fetchParams();
  }

  async function handleDelete(id) {
    await deleteDoc(doc(db, 'auction_parameters', id));
    fetchParams();
  }

  async function handleAdd() {
    await addDoc(collection(db, 'auction_parameters'), form);
    setForm({});
    fetchParams();
  }

  function handleChange(e) {
    setForm({ ...form, [e.target.name]: e.target.value });
  }

  if (loading) return <div>Loading...</div>;

  // Determine which columns to show, in the desired order, and include any extra fields at the end
  const allKeys = params[0] ? Object.keys(params[0]).filter(k => k !== 'id') : [];
  const extraKeys = allKeys.filter(k => !COLUMN_ORDER.includes(k));
  const orderedKeys = [...COLUMN_ORDER, ...extraKeys];

  return (
    <div style={{ padding: 0 }}>
      <div style={{
        overflowX: 'auto',
        overflowY: 'auto',
        maxHeight: '60vh',
        border: '1px solid #eee',
        borderRadius: 8,
        width: '100%',
        boxSizing: 'border-box',
        position: 'relative',
        background: '#fff',
      }}>
        <table border="1" cellPadding="6" style={{ minWidth: 2000, borderCollapse: 'separate', position: 'relative', background: '#fff' }}>
          <thead>
            <tr>
              {orderedKeys.map((key, idx) => (
                <th
                  key={key}
                  style={{
                    position: 'sticky',
                    top: 0,
                    left: idx === 0 ? 0 : undefined,
                    background: idx === 0 ? '#f5f5e6' : '#fafafa',
                    zIndex: idx === 0 ? 4 : 2,
                    borderBottom: '2px solid #ccc',
                  }}
                >
                  {key}
                </th>
              ))}
              <th style={{ position: 'sticky', top: 0, background: '#fafafa', zIndex: 2, borderBottom: '2px solid #ccc' }}>Actions</th>
            </tr>
          </thead>
          <tbody>
            {params.map(param => (
              <tr key={param.id}>
                {orderedKeys.map((key, idx) => (
                  <td
                    key={key}
                    style={idx === 0 ? {
                      position: 'sticky',
                      left: 0,
                      background: '#f9f9e6',
                      zIndex: 3,
                      borderRight: '2px solid #eee',
                      borderLeft: '2px solid #f90', // test border for sticky
                      minWidth: 120,
                      maxWidth: 200,
                    } : {}}
                  >
                    {editing === param.id ? (
                      <input name={key} value={form[key] || ''} onChange={handleChange} />
                    ) : (
                      param[key] || ''
                    )}
                  </td>
                ))}
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
      </div>
      <h3 style={{ marginTop: 24 }}>Add New Parameter</h3>
      <form onSubmit={e => { e.preventDefault(); handleAdd(); }}>
        {orderedKeys.map(key => (
          <input
            key={key}
            name={key}
            placeholder={key}
            value={form[key] || ''}
            onChange={handleChange}
            style={{ marginRight: 8, marginBottom: 8 }}
          />
        ))}
        <button type="submit">Add</button>
      </form>
    </div>
  );
}
