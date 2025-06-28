<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Aplikasi Keuangan Pribadi Sederhana</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://unpkg.com/xlsx/dist/xlsx.full.min.js"></script>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <!-- Chosen Palette: Calm Harmony (Background: #F8F7F4, Containers: #FFFFFF, Text: #3D405B, Income: #A3B18A, Expense: #D98880, Accent: #E07A5F) -->
    <!-- Application Structure Plan: Desain aplikasi ini adalah dasbor satu halaman yang fokus pada kemudahan penggunaan. Struktur utamanya dibagi menjadi dua bagian: panel kontrol di sebelah kiri (atau atas pada perangkat seluler) untuk input transaksi dan kartu ringkasan, serta area tampilan data di sebelah kanan (atau bawah) untuk visualisasi grafik dan riwayat transaksi dalam bentuk tabel. Struktur ini dipilih untuk kesederhanaan maksimal, memungkinkan pengguna untuk dengan cepat mencatat transaksi, melihat gambaran umum keuangan, dan mengekspor data tanpa fitur tambahan yang dapat membingungkan. Ini sangat ideal untuk pengguna yang hanya membutuhkan pelacak pemasukan dan pengeluaran dasar. -->
    <!-- Visualization & Content Choices: Info Laporan -> Tujuan -> Metode Presentasi -> Interaksi -> Justifikasi -> Pustaka/Metode. 1) Input Transaksi -> Mengorganisir -> Formulir HTML -> Pengguna mengisi & mengirim -> Cara paling langsung untuk input data -> HTML & JS. 2) Ringkasan Keuangan -> Menginformasikan -> Kartu Statistik (div) -> Pembaruan dinamis saat data berubah -> Memberikan gambaran cepat kondisi keuangan -> HTML & JS. 3) Pengeluaran per Kategori -> Membandingkan -> Grafik Donat -> Grafik diperbarui, tooltip saat hover -> Visual yang efektif untuk menunjukkan proporsi -> Chart.js (Canvas). 4) Tren Saldo -> Melihat Perubahan -> Grafik Garis -> Grafik diperbarui dengan setiap transaksi -> Cara terbaik untuk menunjukkan tren dari waktu ke waktu -> Chart.js (Canvas). 5) Riwayat Transaksi -> Menginformasikan -> Tabel HTML -> Tabel diperbarui dengan data baru -> Menyediakan catatan terperinci yang dapat diverifikasi -> HTML & JS. 6) Ekspor Data -> Menginformasikan -> Tombol Ekspor -> Mengunduh file XLSX (Excel) -> Memungkinkan pengguna untuk menganalisis data lebih lanjut di luar aplikasi -> HTML & JS, XLSX (SheetJS). Fitur anggaran yang lebih kompleks telah dihapus untuk mencapai kesederhanaan yang diminta. 7) Hapus Transaksi -> Mengorganisir -> Tombol Hapus pada setiap baris riwayat -> Pengguna dapat menghapus transaksi yang salah atau tidak relevan -> Memungkinkan pengelolaan data yang fleksibel -> HTML & JS. -->
    <!-- CONFIRMATION: NO SVG graphics used. NO Mermaid JS used. -->
    <style>
        body {
            font-family: 'Inter', sans-serif;
            background-color: #F8F7F4;
            color: #3D405B;
        }
        .chart-container {
            position: relative;
            width: 100%;
            max-width: 600px;
            margin-left: auto;
            margin-right: auto;
            height: 300px;
            max-height: 350px;
        }
        @media (min-width: 768px) {
            .chart-container {
                height: 320px;
            }
        }
        ::-webkit-scrollbar {
            width: 8px;
        }
        ::-webkit-scrollbar-track {
            background: #F8F7F4;
        }
        ::-webkit-scrollbar-thumb {
            background: #d1d5db;
            border-radius: 10px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #9ca3af;
        }
        .message-box-overlay {
            position: fixed;
            top: 0;
            left: 0;
            width: 100%;
            height: 100%;
            background: rgba(0, 0, 0, 0.5);
            display: flex;
            justify-content: center;
            align-items: center;
            z-index: 1000;
        }
        .message-box-content {
            background: #fff;
            padding: 2rem;
            border-radius: 0.75rem;
            box-shadow: 0 10px 15px rgba(0, 0, 0, 0.1);
            text-align: center;
            max-width: 400px;
            width: 90%;
        }
        .message-box-content button {
            margin-top: 1rem;
            padding: 0.6rem 1.5rem;
            background-color: #E07A5F;
            color: white;
            border: none;
            border-radius: 0.5rem;
            cursor: pointer;
            font-weight: 600;
            transition: background-color 0.2s;
        }
        .message-box-content button:hover {
            background-color: #d16a50;
        }
        .confirm-dialog .message-box-content button:first-of-type {
            background-color: #D98880; /* Red for delete */
            margin-right: 0.5rem;
        }
        .confirm-dialog .message-box-content button:first-of-type:hover {
            background-color: #C67A72;
        }
    </style>
</head>
<body class="antialiased">
    <div class="container mx-auto p-4 md:p-8">
        <header class="text-center mb-8 md:mb-12">
            <h1 class="text-4xl md:text-5xl font-bold text-[#3D405B]">Dasbor Keuangan Pribadi Sederhana</h1>
            <p class="mt-2 text-lg text-gray-600">Lacak pemasukan dan pengeluaran Anda dengan mudah dan fokus.</p>
        </header>

        <main class="grid grid-cols-1 lg:grid-cols-3 gap-8">
            <div class="lg:col-span-1 flex flex-col gap-8">
                <section id="input-section" class="bg-white p-6 rounded-2xl shadow-sm">
                    <h2 class="text-2xl font-bold mb-4">Tambah Transaksi Baru</h2>
                    <p class="mb-6 text-gray-500 text-sm">Isi detail di bawah ini untuk mencatat pemasukan atau pengeluaran baru.</p>
                    <form id="transaction-form" class="space-y-4">
                        <div>
                            <label for="type" class="block text-sm font-medium text-gray-700 mb-2">Jenis Transaksi</label>
                            <div class="grid grid-cols-2 gap-2">
                                <label class="flex items-center justify-center p-3 border rounded-lg cursor-pointer has-[:checked]:bg-[#A3B18A] has-[:checked]:text-white has-[:checked]:border-[#A3B18A] transition-all">
                                    <input type="radio" name="type" value="income" class="sr-only" checked>
                                    <span class="font-semibold">Pemasukan</span>
                                </label>
                                <label class="flex items-center justify-center p-3 border rounded-lg cursor-pointer has-[:checked]:bg-[#D98880] has-[:checked]:text-white has-[:checked]:border-[#D98880] transition-all">
                                    <input type="radio" name="type" value="expense" class="sr-only">
                                    <span class="font-semibold">Pengeluaran</span>
                                </label>
                            </div>
                        </div>
                        <div>
                            <label for="description" class="block text-sm font-medium text-gray-700">Deskripsi</label>
                            <input type="text" id="description" placeholder="Contoh: Gaji bulanan, Belanja" class="mt-1 block w-full px-3 py-2 bg-gray-50 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-[#E07A5F] focus:border-[#E07A5F] sm:text-sm" required>
                        </div>
                        <div>
                            <label for="amount" class="block text-sm font-medium text-gray-700">Jumlah (Rp)</label>
                            <input type="number" id="amount" placeholder="50000" min="0" class="mt-1 block w-full px-3 py-2 bg-gray-50 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-[#E07A5F] focus:border-[#E07A5F] sm:text-sm" required>
                        </div>
                        <div>
                            <label for="category" class="block text-sm font-medium text-gray-700">Kategori</label>
                            <select id="category" class="mt-1 block w-full pl-3 pr-10 py-2 bg-gray-50 text-base border-gray-300 focus:outline-none focus:ring-[#E07A5F] focus:border-[#E07A5F] sm:text-sm rounded-md" required>
                            </select>
                        </div>
                        <div>
                            <label for="date" class="block text-sm font-medium text-gray-700">Tanggal</label>
                            <input type="date" id="date" class="mt-1 block w-full px-3 py-2 bg-gray-50 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-[#E07A5F] focus:border-[#E07A5F] sm:text-sm" required>
                        </div>
                        <button type="submit" class="w-full bg-[#E07A5F] text-white font-bold py-3 px-4 rounded-lg hover:bg-opacity-90 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-[#E07A5F] transition-all shadow">
                            Tambah Transaksi
                        </button>
                    </form>
                </section>

                <section id="summary-section" class="bg-white p-6 rounded-2xl shadow-sm">
                    <h2 class="text-2xl font-bold mb-4">Ringkasan Keuangan</h2>
                    <div class="space-y-4">
                        <div class="flex justify-between items-center p-4 bg-green-50 border-l-4 border-[#A3B18A] rounded-lg">
                            <span class="font-medium">Total Pemasukan</span>
                            <span id="total-income" class="font-bold text-lg text-green-800">Rp 0</span>
                        </div>
                        <div class="flex justify-between items-center p-4 bg-red-50 border-l-4 border-[#D98880] rounded-lg">
                            <span class="font-medium">Total Pengeluaran</span>
                            <span id="total-expense" class="font-bold text-lg text-red-800">Rp 0</span>
                        </div>
                        <div class="flex justify-between items-center p-4 bg-blue-50 border-l-4 border-blue-500 rounded-lg">
                            <span class="font-medium">Saldo Saat Ini</span>
                            <span id="current-balance" class="font-bold text-lg text-blue-800">Rp 0</span>
                        </div>
                    </div>
                    <button id="export-excel-button" class="w-full mt-6 bg-blue-600 text-white font-bold py-3 px-4 rounded-lg hover:bg-blue-700 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-blue-500 transition-all shadow">
                        Ekspor ke Excel (.xlsx)
                    </button>
                </section>
            </div>

            <div class="lg:col-span-2 flex flex-col gap-8">
                <section id="budget-section" class="bg-white p-6 rounded-2xl shadow-sm">
                    <h2 class="text-2xl font-bold mb-4">Anggaran Bulanan</h2>
                    <p class="mb-6 text-gray-500 text-sm">Tetapkan batas pengeluaran untuk setiap kategori dan pantau kemajuan Anda.</p>
                    <form id="budget-form" class="space-y-4">
                        <div id="budget-inputs" class="grid grid-cols-1 md:grid-cols-2 gap-4">
                            <!-- Budget inputs will be dynamically loaded here -->
                        </div>
                        <button type="submit" class="w-full bg-[#A3B18A] text-white font-bold py-3 px-4 rounded-lg hover:bg-opacity-90 focus:outline-none focus:ring-2 focus:ring-offset-2 focus:ring-[#A3B18A] transition-all shadow">
                            Simpan Anggaran
                        </button>
                    </form>
                    <div class="mt-8">
                        <h3 class="text-xl font-semibold mb-4 text-center">Status Anggaran per Kategori</h3>
                        <div class="chart-container">
                            <canvas id="budget-chart"></canvas>
                        </div>
                    </div>
                    <div id="budget-status-list" class="mt-6 space-y-3">
                        <!-- Budget status list will be dynamically loaded here -->
                    </div>
                </section>

                <section id="charts-section" class="bg-white p-6 rounded-2xl shadow-sm">
                     <h2 class="text-2xl font-bold mb-1">Visualisasi Data Transaksi</h2>
                     <p class="mb-6 text-gray-500 text-sm">Analisis keuangan Anda melalui grafik interaktif di bawah ini.</p>
                    <div class="grid grid-cols-1 xl:grid-cols-2 gap-8">
                        <div class="w-full">
                            <h3 class="text-xl font-semibold mb-4 text-center">Alokasi Pengeluaran</h3>
                            <div class="chart-container">
                                <canvas id="expense-chart"></canvas>
                            </div>
                        </div>
                        <div class="w-full">
                             <h3 class="text-xl font-semibold mb-4 text-center">Tren Saldo</h3>
                            <div class="chart-container">
                                <canvas id="balance-chart"></canvas>
                            </div>
                        </div>
                    </div>
                </section>

                <section id="history-section" class="bg-white p-6 rounded-2xl shadow-sm">
                    <h2 class="text-2xl font-bold mb-4">Riwayat Transaksi</h2>
                    <div class="overflow-x-auto max-h-96">
                        <table class="min-w-full divide-y divide-gray-200">
                            <thead class="bg-gray-50 sticky top-0">
                                <tr>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Tanggal</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Deskripsi</th>
                                    <th class="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">Kategori</th>
                                    <th class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Jumlah</th>
                                    <th class="px-6 py-3 text-right text-xs font-medium text-gray-500 uppercase tracking-wider">Aksi</th> <!-- New Column -->
                                </tr>
                            </thead>
                            <tbody id="transaction-table-body" class="bg-white divide-y divide-gray-200">
                                <tr id="no-data-row">
                                    <td colspan="5" class="px-6 py-10 text-center text-gray-500">Belum ada data transaksi.</td> <!-- Colspan updated -->
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </section>
            </div>
        </main>
    </div>

    <script>
        document.addEventListener('DOMContentLoaded', function() {
            const transactionForm = document.getElementById('transaction-form');
            const transactionTypeRadios = document.querySelectorAll('input[name="type"]');
            const categorySelect = document.getElementById('category');
            const dateInput = document.getElementById('date');
            
            const totalIncomeEl = document.getElementById('total-income');
            const totalExpenseEl = document.getElementById('total-expense');
            const currentBalanceEl = document.getElementById('current-balance');
            const transactionTableBody = document.getElementById('transaction-table-body');
            const noDataRow = document.getElementById('no-data-row');
            const exportExcelButton = document.getElementById('export-excel-button');

            const expenseChartCtx = document.getElementById('expense-chart').getContext('2d');
            const balanceChartCtx = document.getElementById('balance-chart').getContext('2d');

            // Budgeting elements
            const budgetForm = document.getElementById('budget-form');
            const budgetInputsContainer = document.getElementById('budget-inputs');
            const budgetStatusList = document.getElementById('budget-status-list');
            const budgetChartCtx = document.getElementById('budget-chart').getContext('2d');

            let transactions = JSON.parse(localStorage.getItem('transactions')) || [];
            let budgets = JSON.parse(localStorage.getItem('budgets')) || {};
            let expenseChart;
            let balanceChart;
            let budgetChart;
            
            const categories = {
                income: ['Gaji', 'Bonus', 'Investasi', 'Hadiah', 'Lain-lain'],
                expense: ['Makan & Minum', 'Transportasi', 'Hiburan', 'Cicilan Rumah', 'Sewa', 'Pendidikan', 'Kesehatan', 'Belanja', 'Lain-lain']
            };

            const formatCurrency = (amount) => {
                return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', minimumFractionDigits: 0 }).format(amount);
            };

            const updateCategoryOptions = (type) => {
                categorySelect.innerHTML = '';
                categories[type].forEach(category => {
                    const option = document.createElement('option');
                    option.value = category;
                    option.textContent = category;
                    categorySelect.appendChild(option);
                });
            };

            const renderTransactions = () => {
                transactionTableBody.innerHTML = '';
                if (transactions.length === 0) {
                    noDataRow.style.display = 'table-row'; // Show no data row
                    return;
                } else {
                    noDataRow.style.display = 'none'; // Hide no data row
                }

                const sortedTransactions = [...transactions].sort((a, b) => new Date(b.date) - new Date(a.date));

                sortedTransactions.forEach(tx => {
                    const row = document.createElement('tr');
                    const isIncome = tx.type === 'income';
                    row.className = isIncome ? 'bg-green-50/30' : 'bg-red-50/30';
                    
                    row.innerHTML = `
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${new Date(tx.date).toLocaleDateString('id-ID')}</td>
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-900">${tx.description}</td>
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-gray-500">${tx.category}</td>
                        <td class="px-6 py-4 whitespace-nowrap text-sm text-right font-medium ${isIncome ? 'text-green-600' : 'text-red-600'}">
                            ${isIncome ? '+' : '-'} ${formatCurrency(tx.amount)}
                        </td>
                        <td class="px-6 py-4 whitespace-nowrap text-right text-sm font-medium">
                            <button data-id="${tx.id}" class="delete-transaction-btn text-red-600 hover:text-red-900 font-bold py-1 px-2 rounded-lg transition-colors">Hapus</button>
                        </td>
                    `;
                    transactionTableBody.appendChild(row);
                });
                // Add event listeners to new delete buttons
                document.querySelectorAll('.delete-transaction-btn').forEach(button => {
                    button.addEventListener('click', function() {
                        const transactionId = parseInt(this.dataset.id);
                        confirmAndDeleteTransaction(transactionId);
                    });
                });
            };

            const updateSummary = () => {
                const totalIncome = transactions.filter(tx => tx.type === 'income').reduce((sum, tx) => sum + tx.amount, 0);
                const totalExpense = transactions.filter(tx => tx.type === 'expense').reduce((sum, tx) => sum + tx.amount, 0);
                const balance = totalIncome - totalExpense;

                totalIncomeEl.textContent = formatCurrency(totalIncome);
                totalExpenseEl.textContent = formatCurrency(totalExpense);
                currentBalanceEl.textContent = formatCurrency(balance);
            };
            
            const createOrUpdateCharts = () => {
                updateExpenseDonutChart();
                updateBalanceLineChart();
                updateBudgetChart();
            };

            const updateExpenseDonutChart = () => {
                const expenseData = transactions
                    .filter(tx => tx.type === 'expense')
                    .reduce((acc, tx) => {
                        if (!acc[tx.category]) {
                            acc[tx.category] = 0;
                        }
                        acc[tx.category] += tx.amount;
                        return acc;
                    }, {});

                const labels = Object.keys(expenseData);
                const data = Object.values(expenseData);

                const chartData = {
                    labels: labels.length > 0 ? labels : ['Tidak ada pengeluaran'],
                    datasets: [{
                        data: data.length > 0 ? data : [1],
                        backgroundColor: [
                            '#F94144', '#F3722C', '#F8961E', '#F9844A', '#F9C74F',
                            '#90BE6D', '#43AA8B', '#4D908E', '#577590', '#277DA1'
                        ],
                         borderColor: '#FFFFFF',
                         borderWidth: 2,
                    }]
                };

                if (expenseChart) {
                    expenseChart.data = chartData;
                    expenseChart.update();
                } else {
                    expenseChart = new Chart(expenseChartCtx, {
                        type: 'doughnut',
                        data: chartData,
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            plugins: {
                                legend: {
                                    position: 'bottom',
                                    labels: {
                                        boxWidth: 12,
                                        padding: 15,
                                    }
                                },
                                tooltip: {
                                    callbacks: {
                                        label: function(context) {
                                            const label = context.label || '';
                                            const value = context.raw || 0;
                                            return `${label}: ${formatCurrency(value)}`;
                                        }
                                    }
                                }
                            },
                            cutout: '60%'
                        }
                    });
                }
            };
            
            const updateBalanceLineChart = () => {
                const sortedTransactions = [...transactions].sort((a,b) => new Date(a.date) - new Date(b.date));
                let runningBalance = 0;
                const balanceData = sortedTransactions.map(tx => {
                    runningBalance += tx.type === 'income' ? tx.amount : -tx.amount;
                    return { x: tx.date, y: runningBalance };
                });
                
                const labels = balanceData.map(d => new Date(d.x).toLocaleDateString('id-ID'));
                const data = balanceData.map(d => d.y);

                const chartData = {
                    labels: labels.length > 0 ? labels : [new Date().toLocaleDateString('id-ID')],
                    datasets: [{
                        label: 'Saldo',
                        data: data.length > 0 ? data : [0],
                        fill: true,
                        backgroundColor: 'rgba(54, 162, 235, 0.2)',
                        borderColor: 'rgba(54, 162, 235, 1)',
                        tension: 0.1,
                        pointBackgroundColor: 'rgba(54, 162, 235, 1)',
                    }]
                };

                if (balanceChart) {
                    balanceChart.data = chartData;
                    balanceChart.update();
                } else {
                    balanceChart = new Chart(balanceChartCtx, {
                        type: 'line',
                        data: chartData,
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                             scales: {
                                y: {
                                    ticks: {
                                        callback: function(value, index, values) {
                                            return formatCurrency(value);
                                        }
                                    }
                                }
                            },
                            plugins: {
                                legend: {
                                    display: false
                                },
                                tooltip: {
                                    callbacks: {
                                        label: function(context) {
                                            const value = context.raw || 0;
                                            return `Saldo: ${formatCurrency(value)}`;
                                        }
                                    }
                                }
                            }
                        }
                    });
                }
            };

            // Budgeting functions
            const renderBudgetInputs = () => {
                budgetInputsContainer.innerHTML = '';
                categories.expense.forEach(category => {
                    const budgetAmount = budgets[category] || 0;
                    const div = document.createElement('div');
                    div.className = 'col-span-1';
                    div.innerHTML = `
                        <label for="budget-${category.replace(/\s/g, '-')}" class="block text-sm font-medium text-gray-700">${category}</label>
                        <input type="number" id="budget-${category.replace(/\s/g, '-')}" name="budget-${category.replace(/\s/g, '-')}" value="${budgetAmount}" min="0" class="mt-1 block w-full px-3 py-2 bg-gray-50 border border-gray-300 rounded-md shadow-sm focus:outline-none focus:ring-[#A3B18A] focus:border-[#A3B18A] sm:text-sm">
                    `;
                    budgetInputsContainer.appendChild(div);
                });
            };

            const calculateMonthlyExpensesByCategory = () => {
                const currentMonth = new Date().getMonth();
                const currentYear = new Date().getFullYear();
                
                return transactions
                    .filter(tx => tx.type === 'expense' && new Date(tx.date).getMonth() === currentMonth && new Date(tx.date).getFullYear() === currentYear)
                    .reduce((acc, tx) => {
                        if (!acc[tx.category]) {
                            acc[tx.category] = 0;
                        }
                        acc[tx.category] += tx.amount;
                        return acc;
                    }, {});
            };

            const updateBudgetChart = () => {
                const actualExpenses = calculateMonthlyExpensesByCategory();
                const budgetCategories = categories.expense;
                
                const labels = budgetCategories;
                const budgetAmounts = budgetCategories.map(cat => budgets[cat] || 0);
                const spentAmounts = budgetCategories.map(cat => actualExpenses[cat] || 0);

                const chartData = {
                    labels: labels,
                    datasets: [
                        {
                            label: 'Anggaran',
                            data: budgetAmounts,
                            backgroundColor: '#A3B18A',
                            borderColor: '#809B66',
                            borderWidth: 1
                        },
                        {
                            label: 'Terpakai',
                            data: spentAmounts,
                            backgroundColor: '#D98880',
                            borderColor: '#C67A72',
                            borderWidth: 1
                        }
                    ]
                };

                if (budgetChart) {
                    budgetChart.data = chartData;
                    budgetChart.update();
                } else {
                    budgetChart = new Chart(budgetChartCtx, {
                        type: 'bar',
                        data: chartData,
                        options: {
                            responsive: true,
                            maintainAspectRatio: false,
                            scales: {
                                y: {
                                    beginAtZero: true,
                                    ticks: {
                                        callback: function(value) {
                                            return formatCurrency(value);
                                        }
                                    }
                                }
                            },
                            plugins: {
                                legend: {
                                    position: 'top',
                                },
                                tooltip: {
                                    callbacks: {
                                        label: function(context) {
                                            const label = context.dataset.label || '';
                                            const value = context.raw || 0;
                                            return `${label}: ${formatCurrency(value)}`;
                                        }
                                    }
                                }
                            }
                        }
                    });
                }
                renderBudgetStatusList();
            };

            const renderBudgetStatusList = () => {
                budgetStatusList.innerHTML = '';
                const actualExpenses = calculateMonthlyExpensesByCategory();
                
                categories.expense.forEach(category => {
                    const budgeted = budgets[category] || 0;
                    const spent = actualExpenses[category] || 0;
                    const remaining = budgeted - spent;
                    
                    const statusClass = remaining >= 0 ? 'text-green-700' : 'text-red-700';
                    const statusText = remaining >= 0 ? `Sisa: ${formatCurrency(remaining)}` : `Lebih: ${formatCurrency(Math.abs(remaining))}`;

                    const item = document.createElement('div');
                    item.className = 'flex justify-between items-center p-3 rounded-lg border bg-gray-50';
                    item.innerHTML = `
                        <span class="font-medium text-gray-800">${category}</span>
                        <div class="text-right">
                            <span class="block text-sm text-gray-600">Anggaran: ${formatCurrency(budgeted)}</span>
                            <span class="block text-sm text-gray-600">Terpakai: ${formatCurrency(spent)}</span>
                            <span class="block text-base font-semibold ${statusClass}">${statusText}</span>
                        </div>
                    `;
                    budgetStatusList.appendChild(item);
                });
            };

            const saveAndRefreshUI = () => {
                localStorage.setItem('transactions', JSON.stringify(transactions));
                localStorage.setItem('budgets', JSON.stringify(budgets));
                updateSummary();
                renderTransactions();
                createOrUpdateCharts();
                renderBudgetInputs();
            };

            const showMessage = (message, isConfirm = false, onConfirm = null) => {
                const messageBoxOverlay = document.createElement('div');
                messageBoxOverlay.className = 'message-box-overlay ' + (isConfirm ? 'confirm-dialog' : '');
                messageBoxOverlay.innerHTML = `
                    <div class="message-box-content">
                        <p class="text-lg font-semibold mb-4 text-[#3D405B]">${message}</p>
                        ${isConfirm ? `
                            <button id="confirm-action" class="bg-red-600 text-white px-6 py-2 rounded-lg hover:bg-red-700 transition-all mr-2">Ya, Hapus</button>
                            <button id="cancel-action" class="bg-gray-400 text-white px-6 py-2 rounded-lg hover:bg-gray-500 transition-all">Batal</button>
                        ` : `
                            <button id="close-message" class="bg-[#E07A5F] text-white px-6 py-2 rounded-lg hover:bg-opacity-90 transition-all">OK</button>
                        `}
                    </div>
                `;
                document.body.appendChild(messageBoxOverlay);

                if (isConfirm) {
                    document.getElementById('confirm-action').addEventListener('click', () => {
                        onConfirm();
                        document.body.removeChild(messageBoxOverlay);
                    });
                    document.getElementById('cancel-action').addEventListener('click', () => {
                        document.body.removeChild(messageBoxOverlay);
                    });
                } else {
                    document.getElementById('close-message').addEventListener('click', () => {
                        document.body.removeChild(messageBoxOverlay);
                    });
                }
            };

            const confirmAndDeleteTransaction = (id) => {
                const transactionToDelete = transactions.find(tx => tx.id === id);
                if (!transactionToDelete) return;

                showMessage(
                    `Apakah Anda yakin ingin menghapus transaksi '${transactionToDelete.description}' (${formatCurrency(transactionToDelete.amount)})?`,
                    true, // isConfirm = true
                    () => { // onConfirm callback
                        transactions = transactions.filter(tx => tx.id !== id);
                        saveAndRefreshUI();
                        showMessage('Transaksi berhasil dihapus.');
                    }
                );
            };

            const exportToXlsx = () => {
                if (transactions.length === 0) {
                    showMessage('Tidak ada transaksi untuk diekspor.');
                    return;
                }

                const transactionHeaders = ["Tanggal", "Deskripsi", "Kategori", "Jenis Transaksi", "Pemasukan (Rp)", "Pengeluaran (Rp)"];
                const transactionData = transactions.map(tx => {
                    const income = tx.type === 'income' ? tx.amount : 0;
                    const expense = tx.type === 'expense' ? tx.amount : 0;
                    return [
                        new Date(tx.date).toLocaleDateString('id-ID'),
                        tx.description,
                        tx.category,
                        tx.type === 'income' ? 'Pemasukan' : 'Pengeluaran',
                        income,
                        expense
                    ];
                });
                const wsTransactions = XLSX.utils.aoa_to_sheet([transactionHeaders, ...transactionData]);

                const totalIncome = transactions.filter(tx => tx.type === 'income').reduce((sum, tx) => sum + tx.amount, 0);
                const totalExpense = transactions.filter(tx => tx.type === 'expense').reduce((sum, tx) => sum + tx.amount, 0);
                const balance = totalIncome - totalExpense;

                const summaryData = [
                    ["Ringkasan Keuangan"],
                    [],
                    ["Total Pemasukan:", totalIncome],
                    ["Total Pengeluaran:", totalExpense],
                    ["Saldo Saat Ini:", balance],
                    [],
                    ["Pengeluaran per Kategori:"],
                ];

                const expenseByCategoryData = transactions
                    .filter(tx => tx.type === 'expense')
                    .reduce((acc, tx) => {
                        if (!acc[tx.category]) {
                            acc[tx.category] = 0;
                        }
                        acc[tx.category] += tx.amount;
                        return acc;
                    }, {});
                
                Object.keys(expenseByCategoryData).forEach(category => {
                    summaryData.push([category, expenseByCategoryData[category]]);
                });

                // Add Budget data to summary sheet
                summaryData.push([]);
                summaryData.push(["Anggaran Bulanan per Kategori:"]);
                categories.expense.forEach(category => {
                    const budgeted = budgets[category] || 0;
                    const spent = expenseByCategoryData[category] || 0; // Actual spent for this month
                    summaryData.push([`Anggaran ${category}:`, budgeted, `Terpakai:`, spent]);
                });


                const wsSummary = XLSX.utils.aoa_to_sheet(summaryData);

                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, wsTransactions, "Transaksi");
                XLSX.utils.book_append_sheet(wb, wsSummary, "Ringkasan");

                XLSX.writeFile(wb, 'data_keuangan_pribadi.xlsx');
            };

            transactionForm.addEventListener('submit', function(e) {
                e.preventDefault();
                const type = document.querySelector('input[name="type"]:checked').value;
                const description = document.getElementById('description').value;
                const amount = parseFloat(document.getElementById('amount').value);
                const category = categorySelect.value;
                const date = dateInput.value;

                if (!description || !amount || !category || !date) {
                    showMessage('Harap isi semua kolom.');
                    return;
                }
                
                transactions.push({ id: Date.now(), type, description, amount, category, date });
                saveAndRefreshUI();
                transactionForm.reset();
                dateInput.valueAsDate = new Date();
                updateCategoryOptions(document.querySelector('input[name="type"]:checked').value);
            });

            transactionTypeRadios.forEach(radio => {
                radio.addEventListener('change', function() {
                    updateCategoryOptions(this.value);
                });
            });

            // Budget form submission
            budgetForm.addEventListener('submit', function(e) {
                e.preventDefault();
                categories.expense.forEach(category => {
                    const inputElement = document.getElementById(`budget-${category.replace(/\s/g, '-')}`);
                    budgets[category] = parseFloat(inputElement.value) || 0;
                });
                saveAndRefreshUI();
                showMessage('Anggaran disimpan!');
            });

            exportExcelButton.addEventListener('click', exportToXlsx);

            const init = () => {
                dateInput.valueAsDate = new Date();
                updateCategoryOptions('income');
                renderBudgetInputs();
                saveAndRefreshUI();
            };
            
            init();
        });
    </script>
</body>
</html>
