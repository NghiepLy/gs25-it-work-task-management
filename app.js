/**
 * IT Work Task Management Dashboard JavaScript Application
 * For BA Manager managing 15 IT Team Members from Jira Excel Data
 * Includes Due Date, Category, and Task Category columns
 */

(function () {
  'use strict';

  // Application State
  let allTasks = [];
  let selectedMember = null; // null means All Members
  let searchQuery = '';
  let monthFilter = '';
  let statusFilter = '';
  let projectFilter = '';
  let categoryFilter = '';
  let taskCatFilter = '';
  let issueTypeFilter = '';
  let sortColumn = 'key';
  let sortDirection = 'asc';
  let currentPage = 1;
  const pageSize = 20;

  // Chart Instances
  let barChartInstance = null;
  let pieChartInstance = null;
  let monthlyTrendChartInstance = null;

  // DOM Elements
  const btnRefresh = document.getElementById('btnRefresh');
  const btnExportHR = document.getElementById('btnExportHR');
  const fileUploadInput = document.getElementById('fileUploadInput');
  const lastUpdatedText = document.getElementById('lastUpdatedText');
  
  const membersGrid = document.getElementById('membersGrid');
  const activeMemberName = document.getElementById('activeMemberName');
  
  const monthSlicerContainer = document.getElementById('monthSlicerContainer');
  const activeMonthName = document.getElementById('activeMonthName');
  
  const kpiTotalTasks = document.getElementById('kpiTotalTasks');
  const kpiCompletedTasks = document.getElementById('kpiCompletedTasks');
  const kpiInProgressTasks = document.getElementById('kpiInProgressTasks');
  const kpiOverdueTasks = document.getElementById('kpiOverdueTasks');
  const kpiCompletionRate = document.getElementById('kpiCompletionRate');

  const searchInput = document.getElementById('searchInput');
  const selectMonth = document.getElementById('filterMonth');
  const selectStatus = document.getElementById('filterStatus');
  const selectProject = document.getElementById('filterProject');
  const selectCategory = document.getElementById('filterCategory');
  const selectTaskCat = document.getElementById('filterTaskCat');
  const selectIssueType = document.getElementById('filterIssueType');
  const btnClearFilters = document.getElementById('btnClearFilters');

  const monthlyMemberTableHead = document.getElementById('monthlyMemberTableHead');
  const monthlyMemberTableBody = document.getElementById('monthlyMemberTableBody');

  const taskTableBody = document.getElementById('taskTableBody');
  const paginationInfo = document.getElementById('paginationInfo');
  const pageIndicator = document.getElementById('pageIndicator');
  const btnPrevPage = document.getElementById('btnPrevPage');
  const btnNextPage = document.getElementById('btnNextPage');
  const toastContainer = document.getElementById('toastContainer');

  // Initialize App
  document.addEventListener('DOMContentLoaded', () => {
    initEventListeners();
    loadData();
  });

  function initEventListeners() {
    // Refresh button click
    btnRefresh.addEventListener('click', refreshDataFromServer);

    // File upload change
    fileUploadInput.addEventListener('change', handleExcelUpload);

    // Export HR Report
    btnExportHR.addEventListener('click', exportHRReport);

    // Search and filter inputs
    searchInput.addEventListener('input', (e) => {
      searchQuery = e.target.value.trim().toLowerCase();
      currentPage = 1;
      renderDashboard();
    });

    selectMonth.addEventListener('change', (e) => {
      monthFilter = e.target.value;
      currentPage = 1;
      renderMonthSlicers();
      renderMemberCards();
      renderDashboard();
    });

    selectStatus.addEventListener('change', (e) => {
      statusFilter = e.target.value;
      currentPage = 1;
      renderDashboard();
    });

    selectProject.addEventListener('change', (e) => {
      projectFilter = e.target.value;
      currentPage = 1;
      renderDashboard();
    });

    selectCategory.addEventListener('change', (e) => {
      categoryFilter = e.target.value;
      currentPage = 1;
      renderDashboard();
    });

    selectTaskCat.addEventListener('change', (e) => {
      taskCatFilter = e.target.value;
      currentPage = 1;
      renderDashboard();
    });

    selectIssueType.addEventListener('change', (e) => {
      issueTypeFilter = e.target.value;
      currentPage = 1;
      renderDashboard();
    });

    btnClearFilters.addEventListener('click', () => {
      searchQuery = '';
      monthFilter = '';
      statusFilter = '';
      projectFilter = '';
      categoryFilter = '';
      taskCatFilter = '';
      issueTypeFilter = '';
      selectedMember = null;
      searchInput.value = '';
      selectMonth.value = '';
      selectStatus.value = '';
      selectProject.value = '';
      selectCategory.value = '';
      selectTaskCat.value = '';
      selectIssueType.value = '';
      currentPage = 1;
      renderMonthSlicers();
      renderMemberCards();
      renderDashboard();
      showToast('Đã xóa tất cả bộ lọc!');
    });

    // Table sorting
    document.querySelectorAll('th.sortable').forEach(th => {
      th.addEventListener('click', () => {
        const col = th.getAttribute('data-sort');
        if (sortColumn === col) {
          sortDirection = sortDirection === 'asc' ? 'desc' : 'asc';
        } else {
          sortColumn = col;
          sortDirection = 'asc';
        }
        renderTable();
      });
    });

    // Pagination
    btnPrevPage.addEventListener('click', () => {
      if (currentPage > 1) {
        currentPage--;
        renderTable();
      }
    });

    btnNextPage.addEventListener('click', () => {
      const maxPages = Math.ceil(getFilteredTasks().length / pageSize);
      if (currentPage < maxPages) {
        currentPage++;
        renderTable();
      }
    });
  }

  // Helper to sanitize tasks array and remove invalid/empty rows
  function sanitizeTasks(tasks) {
    if (!Array.isArray(tasks)) return [];
    return tasks.filter(t => {
      if (!t || typeof t !== 'object') return false;
      const k = String(t.key || '').trim();
      const s = String(t.summary || '').trim();
      if (!k || !s) return false;
      if (k === 'Key' || s === 'Summary' || s === '(No Summary)') return false;
      if (k.startsWith('TASK-') && s === '(No Summary)') return false;
      return true;
    }).map((t, idx) => ({ ...t, id: idx + 1 }));
  }

  // Load Data from data.json or server API
  async function loadData() {
    try {
      showToast('Đang tải dữ liệu tasks IT...', 'info');
      // Try local API first
      let res = await fetch('/api/data').catch(() => null);
      if (!res || !res.ok) {
        // Fallback to static data.json
        res = await fetch('data.json?t=' + Date.now());
      }
      
      if (!res.ok) throw new Error('Không thể tải file dữ liệu');
      const rawTasks = await res.json();
      allTasks = sanitizeTasks(rawTasks);
      
      populateDropdownOptions();
      renderMonthSlicers();
      renderMemberCards();
      renderDashboard();
      
      updateLastUpdatedTime();
      showToast(`Đã tải thành công ${allTasks.length} công việc từ Jira!`, 'success');
    } catch (err) {
      console.error(err);
      showToast('Lỗi khi tải dữ liệu: ' + err.message, 'error');
    }
  }

  // Refresh Data from Server API
  async function refreshDataFromServer() {
    btnRefresh.classList.add('loading');
    showToast('Đang kết nối server đọc lại file IT_Work_Task_Data_JIRA.xlsx...', 'info');
    try {
      let res = await fetch('/api/refresh').catch(() => null);
      if (!res || !res.ok) {
        // Fallback re-fetch data.json
        res = await fetch('data.json?t=' + Date.now());
      }
      if (!res.ok) throw new Error('Không thể cập nhật từ server');
      const rawTasks = await res.json();
      allTasks = sanitizeTasks(rawTasks);
      
      populateDropdownOptions();
      renderMonthSlicers();
      renderMemberCards();
      renderDashboard();
      updateLastUpdatedTime();
      
      showToast(`Cập nhật thành công ${allTasks.length} công việc từ Jira Excel!`, 'success');
    } catch (err) {
      showToast('Không có kết nối server API. Vui lòng chạy server.ps1 hoặc chọn file Excel.', 'error');
    } finally {
      btnRefresh.classList.remove('loading');
    }
  }

  // Handle Client-side Excel File Upload via SheetJS
  function handleExcelUpload(e) {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = function (evt) {
      try {
        const data = new Uint8Array(evt.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        const parsedTasks = jsonRows.map((row, idx) => ({
          id: idx + 1,
          project: row['Project'] || 'Unassigned',
          parent: row['Parent'] || '',
          key: row['Key'] || `TASK-${idx + 1}`,
          issueType: row['Issue Type'] || 'Task',
          summary: row['Summary'] || '(No Summary)',
          assignee: row['Assignee'] || 'Unassigned',
          status: row['Status'] || 'No Status',
          startDate: row['Start date'] || row['Start Date'] || row['StartDate'] || '',
          dueDate: row['Due date'] || row['Due Date'] || '',
          created: row['Created'] || row['Created Date'] || '',
          updated: row['Updated'] || row['Updated Date'] || '',
          dept: row['DEPT'] || '',
          category: row['Category'] || '',
          taskCat: row['Task Category'] || '',
          reviewer: row['Reviewer'] || row['UAT Confirm By'] || '',
          reporter: row['Reporter'] || '',
          timeSpent: row['Time Spent'] || '0',
          creator: row['Creator'] || '',
          subtasks: row['Sub-tasks'] || row['Subtasks'] || ''
        }));

        allTasks = sanitizeTasks(parsedTasks);

        populateDropdownOptions();
        renderMonthSlicers();
        renderMemberCards();
        renderDashboard();
        updateLastUpdatedTime();
        showToast(`Đã nạp thành công ${allTasks.length} tasks từ file ${file.name}!`, 'success');
      } catch (err) {
        showToast('Lỗi khi đọc file Excel: ' + err.message, 'error');
      }
    };
    reader.readAsArrayBuffer(file);
  }

  // Render Month Slicers Grid
  function renderMonthSlicers() {
    if (!monthSlicerContainer) return;

    // Filter tasks by selected member first if any
    const tasksForSlicer = selectedMember 
      ? allTasks.filter(t => t.assignee === selectedMember)
      : allTasks;

    const allMonths = getAllMonthKeys();

    monthSlicerContainer.innerHTML = '';

    // "All Months" slicer card
    const allSlicer = document.createElement('div');
    allSlicer.className = `month-slicer-card ${monthFilter === '' ? 'active' : ''}`;
    allSlicer.innerHTML = `
      <div class="slicer-icon">ALL</div>
      <div class="slicer-details">
        <div class="slicer-month-title">Tất cả các Tháng</div>
        <div class="slicer-task-count">${tasksForSlicer.length} tasks</div>
      </div>
    `;
    allSlicer.addEventListener('click', () => {
      monthFilter = '';
      selectMonth.value = '';
      currentPage = 1;
      renderMonthSlicers();
      renderMemberCards();
      renderDashboard();
    });
    monthSlicerContainer.appendChild(allSlicer);

    // Individual Month slicer cards
    allMonths.forEach(mk => {
      const count = tasksForSlicer.filter(t => isTaskActiveInMonth(t, mk)).length;
      if (count === 0 && monthFilter !== mk) return; // Hide months with 0 active tasks unless selected

      const slicer = document.createElement('div');
      slicer.className = `month-slicer-card ${monthFilter === mk ? 'active' : ''}`;
      
      const monthLabel = formatMonthLabel(mk);
      const monthShort = mk.split('-')[1] ? `T${mk.split('-')[1]}` : 'TH';

      slicer.innerHTML = `
        <div class="slicer-icon month">${monthShort}</div>
        <div class="slicer-details">
          <div class="slicer-month-title">${monthLabel}</div>
          <div class="slicer-task-count">${count} tasks</div>
        </div>
      `;

      slicer.addEventListener('click', () => {
        if (monthFilter === mk) {
          monthFilter = ''; // Deselect if clicked again
          selectMonth.value = '';
        } else {
          monthFilter = mk;
          selectMonth.value = mk;
        }
        currentPage = 1;
        renderMonthSlicers();
        renderMemberCards();
        renderDashboard();
      });

      monthSlicerContainer.appendChild(slicer);
    });

    // Update active month display label
    const activeMonthTag = document.getElementById('activeMonthName');
    if (activeMonthTag) {
      activeMonthTag.textContent = monthFilter ? formatMonthLabel(monthFilter) : 'Tất cả các Tháng (All)';
    }
  }

  // Render Member Selector Cards Grid
  function renderMemberCards() {
    const tasksForMembers = monthFilter
      ? allTasks.filter(t => isTaskActiveInMonth(t, monthFilter))
      : allTasks;

    // Count tasks by assignee
    const counts = {};
    allTasks.forEach(t => {
      const assignee = t.assignee || 'Unassigned';
      if (!counts[assignee]) counts[assignee] = 0;
    });

    tasksForMembers.forEach(t => {
      const assignee = t.assignee || 'Unassigned';
      counts[assignee] = (counts[assignee] || 0) + 1;
    });

    // Sort member names by task count descending
    const memberList = Object.keys(counts).sort((a, b) => counts[b] - counts[a]);

    membersGrid.innerHTML = '';

    // "All Members" card
    const allCard = document.createElement('div');
    allCard.className = `member-card ${selectedMember === null ? 'active' : ''}`;
    allCard.innerHTML = `
      <div class="member-avatar" style="background: linear-gradient(135deg, #6366f1, #8b5cf6);">ALL</div>
      <div class="member-details">
        <div class="member-name">Tất cả Members</div>
        <div class="member-count">${tasksForMembers.length} tasks</div>
      </div>
      <span class="member-badge">${memberList.length} members</span>
    `;
    allCard.addEventListener('click', () => {
      selectedMember = null;
      currentPage = 1;
      renderMemberCards();
      renderMonthSlicers();
      renderDashboard();
    });
    membersGrid.appendChild(allCard);

    // Individual member cards
    memberList.forEach(m => {
      const card = document.createElement('div');
      card.className = `member-card ${selectedMember === m ? 'active' : ''}`;
      
      const initials = getInitials(m);
      
      card.innerHTML = `
        <div class="member-avatar">${initials}</div>
        <div class="member-details">
          <div class="member-name" title="${m}">${m}</div>
          <div class="member-count">${counts[m] || 0} tasks</div>
        </div>
      `;

      card.addEventListener('click', () => {
        if (selectedMember === m) {
          selectedMember = null; // Deselect if clicked again
        } else {
          selectedMember = m;
        }
        currentPage = 1;
        renderMemberCards();
        renderMonthSlicers();
        renderDashboard();
      });

      membersGrid.appendChild(card);
    });

    // Update active member display label
    activeMemberName.textContent = selectedMember ? selectedMember : 'Tất cả Members (All)';
  }

  // Get initials from member name
  function getInitials(name) {
    if (!name || name === 'Unassigned') return 'UN';
    const parts = name.split(/[\s_]+/);
    if (parts.length >= 2) {
      return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
    }
    return name.substring(0, 2).toUpperCase();
  }

  // Helper to parse date string into Date object
  function parseTaskDate(dateStr) {
    if (!dateStr || typeof dateStr !== 'string') return null;
    const str = dateStr.trim();
    if (!str) return null;

    const isoMatch = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (isoMatch) {
      return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
    }

    const dmyMatch = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (dmyMatch) {
      return new Date(parseInt(dmyMatch[3], 10), parseInt(dmyMatch[2], 10) - 1, parseInt(dmyMatch[1], 10));
    }

    return null;
  }

  // Get month range (start and end Date objects) for monthKey "YYYY-MM"
  function getMonthRange(monthKey) {
    if (!monthKey || monthKey === 'Khác') return null;
    const parts = monthKey.split('-');
    if (parts.length !== 2) return null;
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10);

    const start = new Date(year, month - 1, 1, 0, 0, 0);
    const end = new Date(year, month, 0, 23, 59, 59);
    return { start, end, year, month };
  }

  // Enforce 8 Business Rules to determine if a task is active in Month M (monthKey "YYYY-MM")
  function isTaskActiveInMonth(task, monthKey) {
    if (!monthKey) return true;
    if (monthKey === 'Khác') return true;

    const range = getMonthRange(monthKey);
    if (!range) return true;

    const created = parseTaskDate(task.created);
    const start   = parseTaskDate(task.startDate);
    const due     = parseTaskDate(task.dueDate);

    // 1. Task must be initiated on or before Month M end
    let initiated = false;
    if (created && created <= range.end) initiated = true;
    if (start && start <= range.end) initiated = true;
    if (!created && !start) {
      if (due && due >= range.start) initiated = true;
      else {
        const up = parseTaskDate(task.updated);
        if (up && up <= range.end) initiated = true;
      }
    }

    if (!initiated) return false;

    // Current month key (e.g. "2026-07")
    const now = new Date();
    const curYear = now.getFullYear();
    const curMonth = String(now.getMonth() + 1).padStart(2, '0');
    const curMonthKey = `${curYear}-${curMonth}`;

    // 2. Due Date check
    if (due) {
      // Has explicit due date: active if due >= month start
      if (due < range.start) return false;
    } else {
      // NO explicit due date!
      // Do not spill tasks without explicit due dates into future months beyond current month
      if (monthKey > curMonthKey) {
        return false;
      }
    }

    return true;
  }

  // Get all unique months (YYYY-MM) present across created, startDate, dueDate, updated
  function getAllMonthKeys() {
    const keys = new Set();
    allTasks.forEach(t => {
      [t.created, t.startDate, t.dueDate, t.updated].forEach(dStr => {
        const d = parseTaskDate(dStr);
        if (d) {
          const y = d.getFullYear();
          const m = String(d.getMonth() + 1).padStart(2, '0');
          keys.add(`${y}-${m}`);
        }
      });
    });
    return Array.from(keys).sort();
  }

  function formatMonthLabel(monthKey) {
    if (!monthKey || monthKey === 'Khác') return 'Khác';
    const parts = monthKey.split('-');
    if (parts.length === 2) {
      return `Tháng ${parts[1]}/${parts[0]}`;
    }
    return monthKey;
  }

  // Populate dropdown filter select options (Month, Project, Category, Task Category)
  function populateDropdownOptions() {
    // 0. Months
    const months = getAllMonthKeys();
    selectMonth.innerHTML = '<option value="">Tất cả các Tháng</option>';
    months.forEach(m => {
      const opt = document.createElement('option');
      opt.value = m;
      opt.textContent = formatMonthLabel(m);
      selectMonth.appendChild(opt);
    });

    // 1. Projects
    const projects = Array.from(new Set(allTasks.map(t => t.project).filter(Boolean))).sort();
    selectProject.innerHTML = '<option value="">Tất cả Dự án</option>';
    projects.forEach(p => {
      const opt = document.createElement('option');
      opt.value = p;
      opt.textContent = p;
      selectProject.appendChild(opt);
    });

    // 2. Categories
    const categories = Array.from(new Set(allTasks.map(t => t.category).filter(Boolean))).sort();
    selectCategory.innerHTML = '<option value="">Tất cả Category</option>';
    categories.forEach(c => {
      const opt = document.createElement('option');
      opt.value = c;
      opt.textContent = c;
      selectCategory.appendChild(opt);
    });

    // 3. Task Categories
    const taskCats = Array.from(new Set(allTasks.map(t => t.taskCat).filter(Boolean))).sort();
    selectTaskCat.innerHTML = '<option value="">Tất cả Task Category</option>';
    taskCats.forEach(tc => {
      const opt = document.createElement('option');
      opt.value = tc;
      opt.textContent = tc;
      selectTaskCat.appendChild(opt);
    });
  }

  // Filter Tasks based on current selection
  function getFilteredTasks() {
    return allTasks.filter(t => {
      // Member Filter
      if (selectedMember && t.assignee !== selectedMember) return false;

      // Month Filter
      if (monthFilter && !isTaskActiveInMonth(t, monthFilter)) return false;

      // Search Query
      if (searchQuery) {
        const matchKey = t.key.toLowerCase().includes(searchQuery);
        const matchSummary = t.summary.toLowerCase().includes(searchQuery);
        const matchAssignee = t.assignee.toLowerCase().includes(searchQuery);
        const matchProject = t.project.toLowerCase().includes(searchQuery);
        const matchCategory = (t.category || '').toLowerCase().includes(searchQuery);
        const matchTaskCat = (t.taskCat || '').toLowerCase().includes(searchQuery);
        if (!matchKey && !matchSummary && !matchAssignee && !matchProject && !matchCategory && !matchTaskCat) return false;
      }

      // Status Filter
      if (statusFilter) {
        const s = (t.status || '').toLowerCase();
        if (statusFilter === 'To Do' && !s.includes('to do')) return false;
        if (statusFilter === 'In Progress' && !s.includes('progress') && !s.includes('rif.todo')) return false;
        if (statusFilter === 'TESTING' && !s.includes('testing') && !s.includes('review')) return false;
        if (statusFilter === 'Complete' && !s.includes('complete') && !s.includes('done') && !s.includes('release')) return false;
        if (statusFilter === 'Overdue' && !s.includes('overdue')) return false;
      }

      // Project Filter
      if (projectFilter && t.project !== projectFilter) return false;

      // Category Filter
      if (categoryFilter && t.category !== categoryFilter) return false;

      // Task Category Filter
      if (taskCatFilter && t.taskCat !== taskCatFilter) return false;

      // Issue Type Filter
      if (issueTypeFilter && t.issueType !== issueTypeFilter) return false;

      return true;
    });
  }

  // Render Dashboard KPI Metrics, Charts, and Table
  function renderDashboard() {
    const filtered = getFilteredTasks();

    // Calculate KPIs
    const total = filtered.length;
    let completed = 0;
    let inProgress = 0;
    let overdue = 0;

    filtered.forEach(t => {
      const s = (t.status || '').toLowerCase();
      if (s.includes('complete') || s.includes('done') || s.includes('release')) {
        completed++;
      } else if (s.includes('overdue')) {
        overdue++;
      } else {
        inProgress++;
      }
    });

    const completionRate = total > 0 ? Math.round((completed / total) * 100) : 0;

    kpiTotalTasks.textContent = total;
    kpiCompletedTasks.textContent = completed;
    kpiInProgressTasks.textContent = inProgress;
    kpiOverdueTasks.textContent = overdue;
    kpiCompletionRate.textContent = completionRate + '%';

    renderCharts(filtered);
    renderMonthlyMemberTable(filtered);
    renderTable();
  }

  // Render Chart.js Analytics
  function renderCharts(tasks) {
    if (typeof Chart === 'undefined') return;

    // Destroy existing chart instances
    if (barChartInstance) barChartInstance.destroy();
    if (pieChartInstance) pieChartInstance.destroy();
    if (monthlyTrendChartInstance) monthlyTrendChartInstance.destroy();

    // 1. Member Tasks Bar Chart
    const memberStats = {};
    tasks.forEach(t => {
      const assignee = t.assignee || 'Unassigned';
      if (!memberStats[assignee]) {
        memberStats[assignee] = { complete: 0, progress: 0, todo: 0 };
      }
      const s = (t.status || '').toLowerCase();
      if (s.includes('complete') || s.includes('done')) {
        memberStats[assignee].complete++;
      } else if (s.includes('progress') || s.includes('testing')) {
        memberStats[assignee].progress++;
      } else {
        memberStats[assignee].todo++;
      }
    });

    const labels = Object.keys(memberStats).slice(0, 10); // top 10 members
    const completeData = labels.map(l => memberStats[l].complete);
    const progressData = labels.map(l => memberStats[l].progress);
    const todoData = labels.map(l => memberStats[l].todo);

    const ctxBar = document.getElementById('chartMemberTasks').getContext('2d');
    barChartInstance = new Chart(ctxBar, {
      type: 'bar',
      data: {
        labels: labels,
        datasets: [
          { label: 'Complete', data: completeData, backgroundColor: '#10b981' },
          { label: 'In Progress/Testing', data: progressData, backgroundColor: '#f59e0b' },
          { label: 'To Do', data: todoData, backgroundColor: '#3b82f6' }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { stacked: true, ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
        },
        plugins: {
          legend: { labels: { color: '#f8fafc' } }
        }
      }
    });

    // 2. Status Doughnut Chart
    const statusCounts = {};
    tasks.forEach(t => {
      const st = t.status || 'No Status';
      statusCounts[st] = (statusCounts[st] || 0) + 1;
    });

    const pieLabels = Object.keys(statusCounts);
    const pieData = Object.values(statusCounts);
    const pieColors = [
      '#3b82f6', '#10b981', '#f59e0b', '#f43f5e', '#8b5cf6', '#06b6d4', '#64748b'
    ];

    const ctxPie = document.getElementById('chartStatusPie').getContext('2d');
    pieChartInstance = new Chart(ctxPie, {
      type: 'doughnut',
      data: {
        labels: pieLabels,
        datasets: [{
          data: pieData,
          backgroundColor: pieColors,
          borderWidth: 0
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { position: 'right', labels: { color: '#f8fafc', font: { size: 11 } } }
        }
      }
    });

    // 3. Monthly Task Trend Chart
    const allMonths = getAllMonthKeys();
    const trendLabels = [];
    const trendComplete = [];
    const trendInProgress = [];

    allMonths.forEach(mk => {
      const activeInMonth = allTasks.filter(t => isTaskActiveInMonth(t, mk));
      if (activeInMonth.length === 0) return;

      let completeCount = 0;
      let inProgressCount = 0;

      activeInMonth.forEach(t => {
        const s = (t.status || '').toLowerCase();
        if (s.includes('complete') || s.includes('done') || s.includes('release')) {
          completeCount++;
        } else {
          inProgressCount++;
        }
      });

      trendLabels.push(formatMonthLabel(mk));
      trendComplete.push(completeCount);
      trendInProgress.push(inProgressCount);
    });

    const ctxTrend = document.getElementById('chartMonthlyTrend').getContext('2d');
    monthlyTrendChartInstance = new Chart(ctxTrend, {
      type: 'bar',
      data: {
        labels: trendLabels,
        datasets: [
          { label: 'Đã hoàn thành', data: trendComplete, backgroundColor: '#10b981', borderRadius: 4 },
          { label: 'Đang thực hiện / Cần làm', data: trendInProgress, backgroundColor: '#6366f1', borderRadius: 4 }
        ]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        scales: {
          x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
          y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
        },
        plugins: {
          legend: { labels: { color: '#f8fafc' } }
        }
      }
    });
  }

  // Render Monthly Member Matrix Table (Thống kê Tasks theo Member & theo Tháng)
  function renderMonthlyMemberTable(filteredTasks) {
    if (!monthlyMemberTableHead || !monthlyMemberTableBody) return;

    // Get sorted list of months
    const allMonthKeys = getAllMonthKeys();

    // Get list of members from allTasks
    const memberCounts = {};
    allTasks.forEach(t => {
      const m = t.assignee || 'Unassigned';
      memberCounts[m] = (memberCounts[m] || 0) + 1;
    });

    const members = Object.keys(memberCounts).sort((a, b) => memberCounts[b] - memberCounts[a]);

    // Build Matrix Data: matrix[member][monthKey] = active tasks count
    const matrix = {};
    members.forEach(m => {
      matrix[m] = { total: 0, complete: 0 };
      allMonthKeys.forEach(mk => {
        matrix[m][mk] = 0;
      });
    });

    members.forEach(m => {
      const memberTasks = allTasks.filter(t => t.assignee === m);
      memberTasks.forEach(t => {
        allMonthKeys.forEach(mk => {
          if (isTaskActiveInMonth(t, mk)) {
            matrix[m][mk]++;
          }
        });
      });

      const currentFilteredMemberTasks = filteredTasks.filter(t => t.assignee === m);
      matrix[m].total = currentFilteredMemberTasks.length;
      matrix[m].complete = currentFilteredMemberTasks.filter(t => {
        const s = (t.status || '').toLowerCase();
        return s.includes('complete') || s.includes('done') || s.includes('release');
      }).length;
    });

    // Render Table Header
    let headHTML = `
      <tr>
        <th style="text-align: left; width: 180px;">Member / Assignee</th>
    `;
    allMonthKeys.forEach(mk => {
      headHTML += `<th style="text-align: center;">${formatMonthLabel(mk)}</th>`;
    });
    headHTML += `
        <th style="text-align: center;">Tổng Tasks</th>
        <th style="text-align: center;">Hoàn thành (%)</th>
      </tr>
    `;
    monthlyMemberTableHead.innerHTML = headHTML;

    // Render Table Body
    let bodyHTML = '';
    members.forEach(m => {
      const data = matrix[m];
      if (!data || data.total === 0) return; // Skip members with 0 tasks under active filter

      const rate = data.total > 0 ? Math.round((data.complete / data.total) * 100) : 0;
      const rateClass = rate >= 80 ? 'status-complete' : (rate >= 50 ? 'status-progress' : 'status-overdue');

      bodyHTML += `
        <tr>
          <td><strong style="color: var(--text-primary);">${escapeHTML(m)}</strong></td>
      `;
      allMonthKeys.forEach(mk => {
        const cnt = data[mk] || 0;
        const cellStyle = cnt > 0 ? 'font-weight: 600; color: #a5b4fc;' : 'color: var(--text-muted); font-size: 12px;';
        bodyHTML += `<td style="text-align: center; ${cellStyle}">${cnt}</td>`;
      });
      bodyHTML += `
          <td style="text-align: center; font-weight: 700; color: var(--accent-indigo);">${data.total}</td>
          <td style="text-align: center;">
            <span class="badge ${rateClass}" style="font-size: 11px;">${rate}% (${data.complete}/${data.total})</span>
          </td>
        </tr>
      `;
    });

    if (!bodyHTML) {
      bodyHTML = `
        <tr>
          <td colspan="${allMonthKeys.length + 3}" style="text-align: center; padding: 20px; color: var(--text-muted);">
            Không có dữ liệu thống kê phù hợp bộ lọc.
          </td>
        </tr>
      `;
    }

    monthlyMemberTableBody.innerHTML = bodyHTML;
  }

  // Render Task Data Table
  function renderTable() {
    let tasks = getFilteredTasks();

    // Sort tasks
    tasks.sort((a, b) => {
      let valA = a[sortColumn] || '';
      let valB = b[sortColumn] || '';
      if (typeof valA === 'string') valA = valA.toLowerCase();
      if (typeof valB === 'string') valB = valB.toLowerCase();

      if (valA < valB) return sortDirection === 'asc' ? -1 : 1;
      if (valA > valB) return sortDirection === 'asc' ? 1 : -1;
      return 0;
    });

    // Pagination
    const totalCount = tasks.length;
    const maxPages = Math.max(1, Math.ceil(totalCount / pageSize));
    if (currentPage > maxPages) currentPage = maxPages;

    const startIdx = (currentPage - 1) * pageSize;
    const pageTasks = tasks.slice(startIdx, startIdx + pageSize);

    taskTableBody.innerHTML = '';

    if (pageTasks.length === 0) {
      taskTableBody.innerHTML = `
        <tr>
          <td colspan="12" style="text-align: center; padding: 32px; color: var(--text-muted);">
            Không tìm thấy công việc nào phù hợp với bộ lọc.
          </td>
        </tr>
      `;
    } else {
      pageTasks.forEach(t => {
        const tr = document.createElement('tr');
        
        const statusBadge = getStatusBadgeHTML(t.status);
        const issueBadge = getIssueTypeBadgeHTML(t.issueType);
        
        const categoryCell = t.category 
          ? `<span style="font-size: 11.5px; font-weight: 600; color: #a5b4fc; background: rgba(99, 102, 241, 0.12); padding: 3px 8px; border-radius: 6px;">${escapeHTML(t.category)}</span>`
          : `<span style="color: var(--text-muted); font-size: 12px;">-</span>`;

        const taskCatCell = t.taskCat 
          ? `<span style="font-size: 11.5px; font-weight: 600; color: #38bdf8; background: rgba(56, 189, 248, 0.12); padding: 3px 8px; border-radius: 6px;">${escapeHTML(t.taskCat)}</span>`
          : `<span style="color: var(--text-muted); font-size: 12px;">-</span>`;

        const startDateCell = t.startDate
          ? `<span style="font-size: 12px; font-weight: 600; color: #93c5fd; background: rgba(59, 130, 246, 0.15); padding: 3px 8px; border-radius: 6px;">${escapeHTML(t.startDate)}</span>`
          : `<span style="color: var(--text-muted); font-size: 12px;">-</span>`;

        const dueDateCell = t.dueDate
          ? `<span style="font-size: 12px; font-weight: 600; color: #fde68a; background: rgba(245, 158, 11, 0.15); padding: 3px 8px; border-radius: 6px;">${escapeHTML(t.dueDate)}</span>`
          : `<span style="color: var(--text-muted); font-size: 12px;">-</span>`;

        tr.innerHTML = `
          <td><span class="task-key">${escapeHTML(t.key)}</span></td>
          <td>
            <div class="task-summary">${escapeHTML(t.summary)}</div>
            ${t.parent ? `<div class="task-parent">Parent: <strong>${escapeHTML(t.parent)}</strong></div>` : ''}
          </td>
          <td>${issueBadge}</td>
          <td><strong>${escapeHTML(t.assignee)}</strong></td>
          <td><span style="color: var(--text-secondary);">${escapeHTML(t.project)}</span></td>
          <td>${categoryCell}</td>
          <td>${taskCatCell}</td>
          <td>${statusBadge}</td>
          <td>${startDateCell}</td>
          <td>${dueDateCell}</td>
          <td><span style="font-size: 12px; color: var(--text-muted);">${escapeHTML(t.created || '-')}</span></td>
          <td><span style="font-size: 12px; color: var(--text-muted);">${escapeHTML(t.updated || '-')}</span></td>
        `;
        taskTableBody.appendChild(tr);
      });
    }

    // Update Pagination Controls
    paginationInfo.textContent = `Hiển thị ${pageTasks.length ? startIdx + 1 : 0} - ${startIdx + pageTasks.length} / ${totalCount} tasks`;
    pageIndicator.textContent = `Trang ${currentPage} / ${maxPages}`;
    btnPrevPage.disabled = currentPage === 1;
    btnNextPage.disabled = currentPage === maxPages;
  }

  // Status Badge Generator
  function getStatusBadgeHTML(status) {
    const s = (status || '').toLowerCase();
    let cssClass = 'status-todo';
    if (s.includes('complete') || s.includes('done') || s.includes('release')) {
      cssClass = 'status-complete';
    } else if (s.includes('progress') || s.includes('rif.todo')) {
      cssClass = 'status-progress';
    } else if (s.includes('testing') || s.includes('review')) {
      cssClass = 'status-testing';
    } else if (s.includes('overdue')) {
      cssClass = 'status-overdue';
    }
    return `
      <span class="badge ${cssClass}">
        <span class="badge-dot"></span>
        ${escapeHTML(status || 'No Status')}
      </span>
    `;
  }

  // Issue Type Badge Generator
  function getIssueTypeBadgeHTML(issueType) {
    const t = (issueType || '').toLowerCase();
    let cssClass = 'issue-task';
    if (t.includes('sub')) cssClass = 'issue-subtask';
    else if (t.includes('bug')) cssClass = 'issue-bug';
    else if (t.includes('story')) cssClass = 'issue-story';
    else if (t.includes('epic')) cssClass = 'issue-epic';

    return `<span class="badge ${cssClass}">${escapeHTML(issueType || 'Task')}</span>`;
  }

  // Export Filtered Data for HR Report (CSV / Excel)
  function exportHRReport() {
    const tasks = getFilteredTasks();
    if (!tasks.length) {
      showToast('Không có dữ liệu task để xuất báo cáo!', 'error');
      return;
    }

    const headers = ['Key', 'Summary', 'Issue Type', 'Assignee', 'Project', 'Category', 'Task Category', 'Status', 'Start Date', 'Due Date', 'Created Date', 'Updated Date', 'Parent'];
    const rows = tasks.map(t => [
      `"${(t.key || '').replace(/"/g, '""')}"`,
      `"${(t.summary || '').replace(/"/g, '""')}"`,
      `"${(t.issueType || '').replace(/"/g, '""')}"`,
      `"${(t.assignee || '').replace(/"/g, '""')}"`,
      `"${(t.project || '').replace(/"/g, '""')}"`,
      `"${(t.category || '').replace(/"/g, '""')}"`,
      `"${(t.taskCat || '').replace(/"/g, '""')}"`,
      `"${(t.status || '').replace(/"/g, '""')}"`,
      `"${(t.startDate || '').replace(/"/g, '""')}"`,
      `"${(t.dueDate || '').replace(/"/g, '""')}"`,
      `"${(t.created || '').replace(/"/g, '""')}"`,
      `"${(t.updated || '').replace(/"/g, '""')}"`,
      `"${(t.parent || '').replace(/"/g, '""')}"`
    ]);

    const csvContent = '\uFEFF' + [headers.join(','), ...rows.map(r => r.join(','))].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    
    const memberTag = selectedMember ? selectedMember.replace(/\s+/g, '_') : 'All_Members';
    const dateTag = new Date().toISOString().slice(0, 10);
    const fileName = `Bao_Cao_Tasks_HR_IT_${memberTag}_${dateTag}.csv`;

    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', fileName);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);

    showToast(`Đã xuất báo cáo ${fileName} nộp HR thành công!`, 'success');
  }

  // Update Last Updated Timestamp
  function updateLastUpdatedTime() {
    const now = new Date();
    const formatted = now.toLocaleTimeString('vi-VN') + ' ' + now.toLocaleDateString('vi-VN');
    lastUpdatedText.textContent = `Cập nhật: ${formatted}`;
  }

  // Show Toast Notification
  function showToast(message, type = 'info') {
    const toast = document.createElement('div');
    toast.className = `toast toast-${type}`;
    toast.innerHTML = `
      <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><path d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"/></svg>
      <span>${escapeHTML(message)}</span>
    `;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transform = 'translateY(10px)';
      toast.style.transition = 'all 0.3s ease';
      setTimeout(() => toast.remove(), 300);
    }, 4000);
  }

  // Helper HTML escaper
  function escapeHTML(str) {
    if (!str) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

})();
