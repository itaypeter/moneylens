// Same origin — backend serves both API and frontend, so BASE is always empty
const BASE = '';

async function req(method, path, body, isForm = false) {
  const opts = { method, headers: {} };
  if (body && !isForm) {
    opts.headers['Content-Type'] = 'application/json';
    opts.body = JSON.stringify(body);
  } else if (body && isForm) {
    opts.body = body;
  }
  const r = await fetch(`${BASE}/api${path}`, opts);
  if (!r.ok) {
    const err = await r.json().catch(() => ({ error: r.statusText }));
    throw new Error(err.error || r.statusText);
  }
  return r.json();
}

export const api = {
  getTransactions:     ()        => req('GET',    '/transactions'),
  addTransactions:     (txs)     => req('POST',   '/transactions/bulk', { transactions: txs }),
  deleteTransaction:   (id)      => req('DELETE', `/transactions/${id}`),
  updateTxCategory:    (id, cat) => req('PATCH',  `/transactions/${id}`, { category: cat }),
  clearTransactions:   ()        => req('DELETE', '/transactions'),

  getDocuments:        ()        => req('GET',    '/documents'),
  addDocument:         (doc)     => req('POST',   '/documents', doc),
  deleteDocument:      (id)      => req('DELETE', `/documents/${id}`),

  getGroceryReceipts:  ()        => req('GET',    '/grocery-receipts'),
  addGroceryReceipt:   (r)       => req('POST',   '/grocery-receipts', r),
  deleteGroceryReceipt:(id)      => req('DELETE', `/grocery-receipts/${id}`),

  getChildExpenses:    ()        => req('GET',    '/child-expenses'),
  addChildExpense:     (e)       => req('POST',   '/child-expenses', e),
  deleteChildExpense:  (id)      => req('DELETE', `/child-expenses/${id}`),

  getSetting:          (key)     => req('GET',    `/settings/${key}`),
  setSetting:          (key, v)  => req('PUT',    `/settings/${key}`, { value: v }),

  ocrBill: async (file) => {
    const fd = new FormData();
    fd.append('file', file);
    return req('POST', '/ocr/bill', fd, true);
  },
  ocrGrocery: async (file) => {
    const fd = new FormData();
    fd.append('file', file);
    return req('POST', '/ocr/grocery', fd, true);
  },
};
