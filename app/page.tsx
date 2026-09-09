'use client';

import { useState, useEffect } from 'react';
import { TrendingUp, TrendingDown, Wallet, Download, PlusCircle, Search, Check, X, Image as ImageIcon } from 'lucide-react';
import * as XLSX from 'xlsx';

interface Transaction {
  id: number;
  type: 'income' | 'expense';
  amount: number;
  description: string;
  child_name: string | null;
  receipt_url: string | null;
  created_at: string;
}

interface Stats {
  count: number;
  limit: number;
  isActivated: boolean;
  limitType: string;
}

export default function Dashboard() {
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [stats, setStats] = useState<Stats>({ count: 0, limit: 50, isActivated: false, limitType: 'free' });
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState('');
  const [showForm, setShowForm] = useState(false);
  const [showTariffModal, setShowTariffModal] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [selectedFile, setSelectedFile] = useState<File | null>(null);
  const [viewReceiptUrl, setViewReceiptUrl] = useState<string | null>(null);
  const [formData, setFormData] = useState({
    type: 'income',
    amount: '',
    description: '',
    child_name: ''
  });

  useEffect(() => {
    loadData();
  }, []);

  const loadData = async () => {
    try {
      const [transactionsRes, statsRes] = await Promise.all([
        fetch('/api/transactions'),
        fetch('/api/stats')
      ]);
      
      const transactionsData = await transactionsRes.json();
      const statsData = await statsRes.json();
      
      setTransactions(transactionsData);
      setStats(statsData);
    } catch (error) {
      console.error('Error loading data:', error);
    } finally {
      setLoading(false);
    }
  };

  const totalIncome = transactions
    .filter(t => t.type === 'income')
    .reduce((sum, t) => sum + t.amount, 0);
  
  const totalExpense = transactions
    .filter(t => t.type === 'expense')
    .reduce((sum, t) => sum + t.amount, 0);
  
  const balance = totalIncome - totalExpense;

  const filtered = transactions.filter(t => 
    t.description.toLowerCase().includes(search.toLowerCase()) || 
    (t.child_name && t.child_name.toLowerCase().includes(search.toLowerCase()))
  );

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    setUploading(true);
    
    let receiptUrl = null;
    
    if (formData.type === 'expense' && selectedFile) {
      const fileFormData = new FormData();
      fileFormData.append('file', selectedFile);
      
      try {
        const uploadRes = await fetch('/api/upload', {
          method: 'POST',
          body: fileFormData
        });
        
        if (uploadRes.ok) {
          const uploadData = await uploadRes.json();
          receiptUrl = uploadData.url;
        }
      } catch (error) {
        console.error('Upload error:', error);
      }
    }
    
    const res = await fetch('/api/transactions', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        type: formData.type,
        amount: formData.amount,
        description: formData.description,
        child_name: formData.child_name,
        receipt_url: receiptUrl
      })
    });
    
    if (res.status === 403) {
      const errorData = await res.json();
      if (errorData.error === 'LIMIT_REACHED') {
        setShowTariffModal(true);
        setUploading(false);
        return;
      }
    }
    
    if (res.ok) {
      await loadData();
      setShowForm(false);
      setFormData({ type: 'income', amount: '', description: '', child_name: '' });
      setSelectedFile(null);
    } else {
      alert('Ошибка при сохранении');
    }
    
    setUploading(false);
  };

  const handlePayment = (plan: string) => {
    const prices = {
      '200': '750 ₽',
      '500': '1 490 ₽',
      'unlimited': '2 190 ₽'
    };
    
    const names = {
      '200': 'Расширенный (200 записей)',
      '500': 'Профессиональный (500 записей)',
      'unlimited': 'Безлимит'
    };
    
    alert(
      ` Информация по тарифу "${names[plan as keyof typeof names]}"\n\n` +
      `💰 Стоимость: ${prices[plan as keyof typeof prices]}\n\n` +
      `Для активации тарифа свяжитесь с администратором.\n` +
      `Админ-панель: /admin`
    );
  };

  const exportToExcel = () => {
    const data = transactions.map(t => ({
      'ID': t.id,
      'Дата': new Date(t.created_at).toLocaleDateString('ru-RU'),
      'Тип': t.type === 'income' ? 'Доход' : 'Расход',
      'Сумма (₽)': t.amount,
      'Описание': t.description,
      'Ребенок': t.child_name || '-',
      'Чек': t.receipt_url || 'Нет'
    }));

    const ws = XLSX.utils.json_to_sheet(data);
    ws['!cols'] = [
      { wch: 5 },
      { wch: 12 },
      { wch: 10 },
      { wch: 12 },
      { wch: 40 },
      { wch: 20 },
      { wch: 60 }
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Отчет');
    XLSX.writeFile(wb, `otchet_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <div className="text-center">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto mb-4"></div>
          <p className="text-gray-700 font-medium">Загрузка...</p>
        </div>
      </div>
    );
  }

  return (
    <main className="min-h-screen bg-gray-50 p-4 md:p-8">
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <div className="flex flex-col md:flex-row justify-between items-center mb-8 gap-4">
          <div>
            <h1 className="text-3xl font-bold text-gray-900">🏫 Казначей Детского сада</h1>
            <p className="text-gray-700 font-medium mt-1">Учет взносов и расходов группы</p>
          </div>
          <div className="flex gap-2">
            <button 
              onClick={() => setShowForm(!showForm)} 
              className="flex items-center gap-2 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition font-medium"
            >
              <PlusCircle size={18} /> 
              {showForm ? 'Закрыть' : 'Добавить'}
            </button>
            <button 
              onClick={exportToExcel}
              className="flex items-center gap-2 bg-white border-2 border-gray-300 text-gray-800 px-4 py-2 rounded-lg hover:bg-gray-50 transition font-medium"
            >
              <Download size={18} /> 
              Excel
            </button>
          </div>
        </div>

        {/* Stats Cards */}
        <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-8">
          <div className="bg-white p-6 rounded-xl shadow-md border-2 border-gray-200">
            <div className="flex items-center gap-3 text-green-700 mb-2">
              <TrendingUp size={24} />
              <span className="text-base font-bold">Собрано</span>
            </div>
            <p className="text-3xl font-bold text-gray-900">{totalIncome.toLocaleString()} ₽</p>
          </div>
          <div className="bg-white p-6 rounded-xl shadow-md border-2 border-gray-200">
            <div className="flex items-center gap-3 text-red-700 mb-2">
              <TrendingDown size={24} />
              <span className="text-base font-bold">Потрачено</span>
            </div>
            <p className="text-3xl font-bold text-gray-900">{totalExpense.toLocaleString()} ₽</p>
          </div>
          <div className="bg-white p-6 rounded-xl shadow-md border-2 border-gray-200">
            <div className="flex items-center gap-3 text-blue-700 mb-2">
              <Wallet size={24} />
              <span className="text-base font-bold">Остаток</span>
            </div>
            <p className="text-3xl font-bold text-gray-900">{balance.toLocaleString()} ₽</p>
          </div>
        </div>

        {/* Progress Bar - Only show if NOT activated */}
        {!stats.isActivated && (
          <div className="bg-white rounded-xl shadow-md border-2 border-gray-200 p-6 mb-8">
            <div className="flex items-center justify-between mb-3">
              <div>
                <h3 className="text-lg font-bold text-gray-900">📊 Использование лимита</h3>
                <p className="text-sm text-gray-600 font-medium">
                  {stats.count} из {stats.limit} записей
                </p>
              </div>
              <button
                onClick={() => setShowTariffModal(true)}
                className="bg-gradient-to-r from-orange-500 to-amber-500 text-white px-4 py-2 rounded-lg font-bold hover:from-orange-600 hover:to-amber-600 transition flex items-center gap-2"
              >
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <polygon points="13 2 3 14 12 14 11 22 21 10 12 10 13 2"></polygon>
                </svg>
                Увеличить лимит
              </button>
            </div>
            <div className="w-full bg-gray-200 rounded-full h-3">
              <div
                className="bg-gradient-to-r from-green-500 to-emerald-500 h-3 rounded-full transition-all duration-500"
                style={{ width: `${Math.min((stats.count / stats.limit) * 100, 100)}%` }}
              ></div>
            </div>
            <p className="text-sm text-gray-600 font-medium mt-2">
              Осталось: {Math.max(stats.limit - stats.count, 0)} записей
            </p>
          </div>
        )}

        {/* Activated Tariff Block */}
        {stats.isActivated && (
          <div className="bg-gradient-to-r from-green-500 to-emerald-600 rounded-xl shadow-md p-6 mb-8 text-white">
            <div className="flex items-center gap-3">
              <Check size={32} />
              <div>
                {stats.limit === 999999 || stats.limitType === 'unlimited' ? (
                  <p className="font-bold text-xl">✅ Безлимит активирован!</p>
                ) : (
                  <>
                    <p className="font-bold text-xl">✅ Тариф на {stats.limit} записей активирован!</p>
                    <p className="font-medium opacity-90">Использовано: {stats.count} из {stats.limit}</p>
                  </>
                )}
              </div>
            </div>
          </div>
        )}

        {/* Add Form */}
        {showForm && (
          <div className="bg-white p-6 rounded-xl shadow-md border-2 border-gray-200 mb-8">
            <h2 className="text-xl font-bold mb-4 text-gray-900">Новая операция</h2>
            <form onSubmit={handleSubmit} className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <select 
                value={formData.type} 
                onChange={e => setFormData({...formData, type: e.target.value})} 
                className="p-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none font-medium text-gray-900"
              >
                <option value="income">💰 Взнос (Доход)</option>
                <option value="expense">💸 Расход</option>
              </select>
              <input 
                type="number" 
                placeholder="Сумма" 
                value={formData.amount} 
                onChange={e => setFormData({...formData, amount: e.target.value})} 
                className="p-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none font-medium text-gray-900 placeholder-gray-500"
                required 
              />
              <input 
                type="text" 
                placeholder="На что / Кто сдал" 
                value={formData.description} 
                onChange={e => setFormData({...formData, description: e.target.value})} 
                className="p-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none font-medium text-gray-900 placeholder-gray-500"
                required 
              />
              <input 
                type="text" 
                placeholder="Имя ребенка" 
                value={formData.child_name} 
                onChange={e => setFormData({...formData, child_name: e.target.value})} 
                className="p-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 outline-none font-medium text-gray-900 placeholder-gray-500"
              />
              
              {/* Receipt Upload - Only for expenses */}
              {formData.type === 'expense' && (
                <div className="md:col-span-4">
                  <label className="block text-sm font-bold text-gray-900 mb-2">
                    📎 Чек (фото, необязательно)
                  </label>
                  <input 
                    type="file" 
                    accept="image/*"
                    onChange={(e) => setSelectedFile(e.target.files?.[0] || null)}
                    className="w-full p-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 font-medium"
                  />
                  {selectedFile && (
                    <p className="text-sm text-green-700 font-medium mt-2">
                      ✅ Выбран файл: {selectedFile.name}
                    </p>
                  )}
                  <p className="text-xs text-gray-500 mt-1">
                    Поддерживаются: JPG, PNG, WEBP (макс 32 МБ)
                  </p>
                </div>
              )}

              <button 
                type="submit" 
                disabled={uploading}
                className="md:col-span-4 bg-gray-900 text-white py-3 rounded-lg hover:bg-gray-800 transition font-bold disabled:opacity-50"
              >
                {uploading ? 'Загрузка...' : 'Сохранить операцию'}
              </button>
            </form>
          </div>
        )}

        {/* Transactions Table */}
        <div className="bg-white rounded-xl shadow-md border-2 border-gray-200 overflow-hidden">
          <div className="p-4 border-b-2 flex gap-2 bg-gray-50">
            <Search size={18} className="text-gray-600 mt-2" />
            <input 
              placeholder="Поиск по описанию или имени ребенка..." 
              value={search} 
              onChange={e => setSearch(e.target.value)}
              className="w-full p-2 outline-none font-medium text-gray-900 placeholder-gray-600"
            />
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-left">
              <thead className="bg-gray-100 text-gray-900 text-sm">
                <tr>
                  <th className="p-4 font-bold">Дата</th>
                  <th className="p-4 font-bold">Тип</th>
                  <th className="p-4 font-bold">Описание</th>
                  <th className="p-4 font-bold">Ребенок</th>
                  <th className="p-4 font-bold">Чек</th>
                  <th className="p-4 text-right font-bold">Сумма</th>
                </tr>
              </thead>
              <tbody>
                {filtered.length === 0 ? (
                  <tr>
                    <td colSpan={6} className="p-8 text-center text-gray-600 font-medium">
                      Нет операций
                    </td>
                  </tr>
                ) : (
                  filtered.map(t => (
                    <tr key={t.id} className="border-t-2 hover:bg-gray-50 transition">
                      <td className="p-4 text-gray-900 font-medium">
                        {new Date(t.created_at).toLocaleDateString('ru-RU')}
                      </td>
                      <td className="p-4">
                        <span className={`px-3 py-1 rounded-full text-xs font-bold ${
                          t.type === 'income' 
                            ? 'bg-green-100 text-green-800' 
                            : 'bg-red-100 text-red-800'
                        }`}>
                          {t.type === 'income' ? 'Доход' : 'Расход'}
                        </span>
                      </td>
                      <td className="p-4 text-gray-900 font-medium">{t.description}</td>
                      <td className="p-4 text-gray-800 font-medium">{t.child_name || '-'}</td>
                      <td className="p-4 text-center">
                        {t.receipt_url ? (
                          <button
                            onClick={() => setViewReceiptUrl(t.receipt_url)}
                            className="text-blue-600 hover:text-blue-800 font-medium flex items-center gap-1 mx-auto"
                          >
                            <ImageIcon size={16} />
                            <span className="text-sm">Просмотр</span>
                          </button>
                        ) : (
                          <span className="text-gray-400">-</span>
                        )}
                      </td>
                      <td className={`p-4 text-right font-bold text-lg ${
                        t.type === 'income' ? 'text-green-700' : 'text-red-700'
                      }`}>
                        {t.type === 'income' ? '+' : '-'}{t.amount.toLocaleString()} ₽
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>

        {/* Receipt Modal */}
        {viewReceiptUrl && (
          <div className="fixed inset-0 bg-black bg-opacity-75 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-xl max-w-4xl w-full max-h-[90vh] overflow-auto">
              <div className="flex justify-between items-center p-4 border-b-2">
                <h3 className="text-xl font-bold text-gray-900">📎 Чек</h3>
                <button 
                  onClick={() => setViewReceiptUrl(null)}
                  className="p-2 hover:bg-gray-100 rounded-lg transition"
                >
                  <X size={24} />
                </button>
              </div>
              <div className="p-4">
                <img 
                  src={viewReceiptUrl} 
                  alt="Чек" 
                  className="w-full h-auto rounded-lg"
                />
              </div>
              <div className="p-4 border-t-2 flex justify-end gap-2">
                <a 
                  href={viewReceiptUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="px-4 py-2 bg-blue-600 text-white rounded-lg font-medium hover:bg-blue-700 transition"
                >
                  Открыть в новой вкладке
                </a>
                <button 
                  onClick={() => setViewReceiptUrl(null)}
                  className="px-4 py-2 bg-gray-200 text-gray-800 rounded-lg font-medium hover:bg-gray-300 transition"
                >
                  Закрыть
                </button>
              </div>
            </div>
          </div>
        )}

        {/* Tariff Modal */}
        {showTariffModal && (
          <div className="fixed inset-0 bg-black bg-opacity-75 flex items-center justify-center z-50 p-4">
            <div className="bg-white rounded-xl max-w-5xl w-full max-h-[90vh] overflow-auto">
              <div className="flex justify-between items-center p-6 border-b-2">
                <h2 className="text-2xl font-bold text-gray-900 flex items-center gap-2">
                  <span className="text-3xl">🚀</span> Выберите тариф
                </h2>
                <button
                  onClick={() => setShowTariffModal(false)}
                  className="p-2 hover:bg-gray-100 rounded-lg transition"
                >
                  <X size={24} />
                </button>
              </div>
              <div className="p-6 grid grid-cols-1 md:grid-cols-3 gap-6">
                {/* Расширенный */}
                <div className="border-2 border-blue-200 rounded-xl p-6 bg-blue-50 hover:shadow-lg transition">
                  <div className="flex items-center gap-2 mb-4">
                    <span className="text-2xl"></span>
                    <h3 className="text-xl font-bold text-gray-900">Расширенный</h3>
                  </div>
                  <p className="text-3xl font-bold text-gray-900 mb-2">750 ₽</p>
                  <p className="text-gray-600 mb-4">200 записей</p>
                  <ul className="space-y-2 mb-6">
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>200 записей</span>
                    </li>
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>Экспорт в Excel</span>
                    </li>
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>Поиск по записям</span>
                    </li>
                  </ul>
                  <button
                    onClick={() => handlePayment('200')}
                    className="w-full bg-blue-600 text-white py-3 rounded-lg font-bold hover:bg-blue-700 transition"
                  >
                    Выбрать
                  </button>
                </div>

                {/* Профессиональный */}
                <div className="border-2 border-purple-200 rounded-xl p-6 bg-purple-50 hover:shadow-lg transition">
                  <div className="flex items-center gap-2 mb-4">
                    <span className="text-2xl">⭐</span>
                    <h3 className="text-xl font-bold text-gray-900">Профессиональный</h3>
                  </div>
                  <p className="text-3xl font-bold text-gray-900 mb-2">1 490 ₽</p>
                  <p className="text-gray-600 mb-4">500 записей</p>
                  <ul className="space-y-2 mb-6">
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>500 записей</span>
                    </li>
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>Экспорт в Excel</span>
                    </li>
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>Поиск по записям</span>
                    </li>
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>Приоритетная поддержка</span>
                    </li>
                  </ul>
                  <button
                    onClick={() => handlePayment('500')}
                    className="w-full bg-purple-600 text-white py-3 rounded-lg font-bold hover:bg-purple-700 transition"
                  >
                    Выбрать
                  </button>
                </div>

                {/* Безлимит */}
                <div className="border-2 border-yellow-300 rounded-xl p-6 bg-yellow-50 hover:shadow-lg transition">
                  <div className="flex items-center gap-2 mb-4">
                    <span className="text-2xl">👑</span>
                    <h3 className="text-xl font-bold text-gray-900">Безлимит</h3>
                  </div>
                  <p className="text-3xl font-bold text-gray-900 mb-2">2 190 ₽</p>
                  <p className="text-gray-600 mb-4">Безлимит</p>
                  <ul className="space-y-2 mb-6">
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>Безлимитные записи</span>
                    </li>
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>Все функции</span>
                    </li>
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>Пожизненный доступ</span>
                    </li>
                    <li className="flex items-center gap-2 text-gray-700">
                      <Check size={16} className="text-green-600" />
                      <span>VIP поддержка</span>
                    </li>
                  </ul>
                  <button
                    onClick={() => handlePayment('unlimited')}
                    className="w-full bg-yellow-600 text-white py-3 rounded-lg font-bold hover:bg-yellow-700 transition"
                  >
                    Выбрать
                  </button>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    </main>
  );
}