// 全局变量
let projectData = [];
let currentPage = 'home';
let personStatsSort = { key: null, asc: false };
// 日历视图当前月份
let calendarViewMonth = null;
// 新增：全局排班数据
let dailyScheduleData = {};

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    // 初始化导航
    initNavigation();
    // 加载项目数据
    loadProjectData();
    // 初始化事件监听器
    initEventListeners();
    initPersonStatsSort();
    initDailySchedulePage();
    const exportBtn = document.getElementById('exportDailySchedule');
    if (exportBtn) {
        exportBtn.onclick = exportDailySchedule;
    }
    const showHistoryBtn = document.getElementById('showDailyHistory');
    if (showHistoryBtn) {
        showHistoryBtn.onclick = function() {
            renderDailyHistoryList();
            const modal = new bootstrap.Modal(document.getElementById('dailyHistoryModal'));
            modal.show();
        };
    }
});

// 初始化导航
function initNavigation() {
    const navLinks = document.querySelectorAll('.nav-link');
    navLinks.forEach(link => {
        link.addEventListener('click', function(e) {
            e.preventDefault();
            const targetPage = this.getAttribute('data-page');
            showPage(targetPage);
        });
    });
}

// 显示指定页面
function showPage(pageId) {
    // 隐藏所有页面
    document.querySelectorAll('.page-content').forEach(page => {
        page.classList.add('d-none');
    });
    // 显示目标页面
    document.getElementById(pageId).classList.remove('d-none');
    // 更新导航栏激活状态
    document.querySelectorAll('.nav-link').forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('data-page') === pageId) {
            link.classList.add('active');
        }
    });
    // 更新当前页面
    currentPage = pageId;
    // 根据页面类型执行特定操作
    switch(pageId) {
        case 'home':
            updateDashboard();
            break;
        case 'project-distribution':
            updateProjectList();
            break;
        case 'personal-stats':
            updatePersonStats();
            break;
        case 'holiday-schedule':
            updateScheduleRules();
            break;
    }
}

// 加载项目数据
function loadProjectData() {
    // 优先从localStorage读取
    const local = localStorage.getItem('projectData');
    if (local) {
        try {
            projectData = JSON.parse(local);
            updateDashboard();
            return;
        } catch (e) {
            // 解析失败则清空
            localStorage.removeItem('projectData');
        }
    }
    // 显示文件上传提示
    const homeContent = document.getElementById('home');
    homeContent.innerHTML = `
        <div class="alert alert-info" role="alert">
            <h4 class="alert-heading">欢迎使用项目人员信息管理系统</h4>
            <p>请先上传项目信息表以开始使用系统。</p>
            <hr>
            <div class="input-group">
                <input type="file" class="form-control" id="initialProjectInfoFile" accept=".xlsx,.xls">
                <button class="btn btn-primary" id="initialUploadBtn">上传项目信息表</button>
            </div>
        </div>
    `;
    // 添加文件上传事件监听
    document.getElementById('initialUploadBtn').addEventListener('click', function() {
        document.getElementById('initialProjectInfoFile').click();
    });
    document.getElementById('initialProjectInfoFile').addEventListener('change', function(event) {
        const file = event.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                projectData = XLSX.utils.sheet_to_json(firstSheet, {defval: ''});
                // 保存到localStorage
                localStorage.setItem('projectData', JSON.stringify(projectData));
                // 恢复首页显示
                homeContent.innerHTML = `
                    <div class="row">
                        <div class="col-md-4">
                            <div class="card">
                                <div class="card-body">
                                    <h5 class="card-title">项目总数</h5>
                                    <p class="card-text" id="totalProjects">加载中...</p>
                                </div>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="card">
                                <div class="card-body">
                                    <h5 class="card-title">服务经理总数</h5>
                                    <p class="card-text" id="totalManagers">加载中...</p>
                                </div>
                            </div>
                        </div>
                        <div class="col-md-4">
                            <div class="card">
                                <div class="card-body">
                                    <h5 class="card-title">工程师总数</h5>
                                    <p class="card-text" id="totalEngineers">加载中...</p>
                                </div>
                            </div>
                        </div>
                    </div>
                `;
                updateDashboard();
                if (currentPage === 'project-distribution') {
                    updateProjectList();
                } else if (currentPage === 'personal-stats') {
                    updatePersonStats();
                }
            };
            reader.readAsArrayBuffer(file);
        }
    });
}

// 更新仪表盘数据
function updateDashboard() {
    document.getElementById('totalProjects').textContent = projectData.length;
    
    // 获取所有服务经理
    const managers = new Set(projectData.map(p => p['服务经理']));
    
    // 获取所有工程师
    const engineers = new Set();
    projectData.forEach(project => {
        // 遍历所有可能的工程师列
        for (let i = 1; ; i++) {
            const engineerKey = `工程师${i}`;
            if (project[engineerKey] && project[engineerKey].trim() !== '') {
                engineers.add(project[engineerKey]);
            } else {
                break;
            }
        }
    });

    document.getElementById('totalManagers').textContent = managers.size;
    document.getElementById('totalEngineers').textContent = engineers.size;
}

// 更新项目列表
function updateProjectList() {
    const projectList = document.getElementById('projectList');
    projectList.innerHTML = '';
    projectData.forEach(project => {
        const item = document.createElement('a');
        item.href = '#';
        item.className = 'list-group-item list-group-item-action';
        item.textContent = project['项目'];
        item.addEventListener('click', () => showProjectDetails(project));
        projectList.appendChild(item);
    });
}

// 显示项目详情
function showProjectDetails(project) {
    const details = document.getElementById('projectDetails');
    let engineersHtml = '';
    
    // 遍历所有工程师
    for (let i = 1; ; i++) {
        const engineerKey = `工程师${i}`;
        if (project[engineerKey] && project[engineerKey].trim() !== '') {
            engineersHtml += `<p><strong>工程师${i}：</strong>${project[engineerKey]}</p>`;
        } else {
            break;
        }
    }

    details.innerHTML = `
        <h4>${project['项目']}</h4>
        <div class="mt-3">
            <p><strong>项目分值：</strong>${project['项目分值'] || ''}</p>
            <p><strong>服务经理：</strong>${project['服务经理']}</p>
            <p><strong>SM电话号码：</strong>${project['SM电话号码'] || ''}</p>
            <p><strong>项目邮箱组：</strong>${project['项目邮箱组'] || ''}</p>
            <div class="mt-2">
                <strong>工程师团队：</strong>
                ${engineersHtml}
            </div>
        </div>
    `;
}

// 更新人员统计
function updatePersonStats() {
    const personType = document.getElementById('personType').value;
    const searchText = document.getElementById('personSearch').value.toLowerCase();
    const table = document.getElementById('personStatsTable');
    table.innerHTML = '';

    const personStats = {};
    
    projectData.forEach(project => {
        if (personType === 'manager') {
            const manager = project['服务经理'];
            if (manager.toLowerCase().includes(searchText)) {
                if (!personStats[manager]) {
                    personStats[manager] = {
                        type: '服务经理',
                        projects: [],
                        scores: []
                    };
                }
                personStats[manager].projects.push(project['项目']);
                let score = Number(project['项目分值']) || 0;
                personStats[manager].scores.push(score);
            }
        } else {
            for (let i = 1; ; i++) {
                const engineerKey = `工程师${i}`;
                if (project[engineerKey] && project[engineerKey].trim() !== '') {
                    const engineer = project[engineerKey];
                    if (engineer.toLowerCase().includes(searchText)) {
                        if (!personStats[engineer]) {
                            personStats[engineer] = {
                                type: '工程师',
                                projects: [],
                                scores: []
                            };
                        }
                        personStats[engineer].projects.push(project['项目']);
                        let score = Number(project['项目分值']) || 0;
                        personStats[engineer].scores.push(score);
                    }
                } else {
                    break;
                }
            }
        }
    });

    // 排序处理
    let statsArr = Object.entries(personStats).map(([name, data]) => {
        return {
            name,
            type: data.type,
            projects: data.projects,
            projectCount: data.projects.length,
            scoreSum: data.scores.reduce((a, b) => a + b, 0)
        };
    });
    if (personStatsSort.key === 'projectCount') {
        statsArr.sort((a, b) => {
            if (a.projectCount === b.projectCount) {
                // 项目数量相同，按分值总和排序
                if (personStatsSort.asc) {
                    return a.scoreSum - b.scoreSum;
                } else {
                    return b.scoreSum - a.scoreSum;
                }
            }
            if (personStatsSort.asc) {
                return a.projectCount - b.projectCount;
            } else {
                return b.projectCount - a.projectCount;
            }
        });
    }

    statsArr.forEach(data => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${data.name}</td>
            <td>${data.type}</td>
            <td>${data.projects.join(', ')}</td>
            <td>${data.projectCount}</td>
            <td>${data.scoreSum}</td>
        `;
        table.appendChild(row);
    });
}

// 只保留项目数量排序
function initPersonStatsSort() {
    const projectCountTh = document.getElementById('sortProjectCount');
    const projectCountIcon = document.getElementById('projectCountSortIcon');
    if (projectCountTh) {
        projectCountTh.onclick = function() {
            if (personStatsSort.key === 'projectCount') {
                personStatsSort.asc = !personStatsSort.asc;
            } else {
                personStatsSort.key = 'projectCount';
                personStatsSort.asc = false;
            }
            projectCountIcon.textContent = personStatsSort.asc ? '▲' : '▼';
            updatePersonStats();
        };
    }
}

// 更新排班规则
function updateScheduleRules() {
    const rules = document.getElementById('scheduleRules');
    rules.innerHTML = `
        <h6>排班规则：</h6>
        <ul>
            <li>同一项目的所有工程师尽量轮流分担值班任务</li>
            <li>同一工程师可以同时负责多个项目</li>
            <li>同一工程师的值班任务尽量集中安排在排班周期的连续几天</li>
            <li>每个项目必须配备一名服务经理和一名工程师值班</li>
        </ul>
    `;
}

// 动态填充项目筛选下拉框
function updateProjectFilter() {
    const filter = document.getElementById('projectFilter');
    if (!filter) return;
    // 清空原有选项，保留"全部项目"
    filter.innerHTML = '<option value="all">全部项目</option>';
    const names = [...new Set(projectData.map(p => p['项目']))];
    names.forEach(name => {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        filter.appendChild(opt);
    });
}

// 生成排班表（支持筛选和折叠/展开）
function generateSchedule() {
    const container = document.querySelector('#scheduleTable').parentElement;
    // 清空原有内容
    container.innerHTML = '<table class="table table-striped" id="scheduleTable"></table>';
    // 获取筛选条件
    const filter = document.getElementById('projectFilter');
    const filterValue = filter ? filter.value : 'all';
    // 获取用户选择的日期
    const startDateInput = document.getElementById('startDate').value;
    const endDateInput = document.getElementById('endDate').value;
    if (!startDateInput || !endDateInput) {
        showCustomAlert('请选择开始日期和结束日期！', '错误');
        return;
    }
    const startDate = new Date(startDateInput);
    const endDate = new Date(endDateInput);
    if (startDate > endDate) {
        showCustomAlert('开始日期不能晚于结束日期！', '错误');
        return;
    }
    // 生成日期范围内的所有日期
    const dates = [];
    const currentDate = new Date(startDate);
    while (currentDate <= endDate) {
        dates.push(currentDate.toISOString().split('T')[0]);
        currentDate.setDate(currentDate.getDate() + 1);
    }
    // 过滤项目
    let showProjects = projectData;
    if (filterValue && filterValue !== 'all') {
        showProjects = projectData.filter(p => p['项目'] === filterValue);
    }
    // 添加排班周期标题
    const scheduleTitle = document.createElement('h5');
    scheduleTitle.className = 'mb-3';
    scheduleTitle.textContent = `排班表 (${startDate.toLocaleDateString()} - ${endDate.toLocaleDateString()})`;
    container.appendChild(scheduleTitle);
    // 生成每个项目的折叠卡片
    showProjects.forEach((project, idx) => {
        // 卡片外层
        const card = document.createElement('div');
        card.className = 'card mb-3';
        // 卡片头部（可点击折叠/展开）
        const cardHeader = document.createElement('div');
        cardHeader.className = 'card-header';
        cardHeader.style.cursor = 'pointer';
        cardHeader.setAttribute('data-bs-toggle', 'collapse');
        const collapseId = `collapseProject${idx}`;
        cardHeader.setAttribute('data-bs-target', `#${collapseId}`);
        cardHeader.innerHTML = `<strong>项目：${project['项目']}（服务经理：${project['服务经理']}）</strong>`;
        card.appendChild(cardHeader);
        // 卡片内容（默认收起）
        const cardBody = document.createElement('div');
        cardBody.className = 'collapse';
        cardBody.id = collapseId;
        // 展示三项信息
        const infoDiv = document.createElement('div');
        infoDiv.className = 'mb-2';
        infoDiv.innerHTML = `
            <span><strong>SM电话号码：</strong>${project['SM电话号码'] || ''}</span>&nbsp;&nbsp;
            <span><strong>项目邮箱组：</strong>${project['项目邮箱组'] || ''}</span>&nbsp;&nbsp;
            <span><strong>Call Center联系方式：</strong>${project['Call Center联系方式'] || ''}</span>
        `;
        cardBody.appendChild(infoDiv);
        // 获取所有工程师
        const engineers = [];
        for (let i = 1; ; i++) {
            const engineerKey = `工程师${i}`;
            if (project[engineerKey] && project[engineerKey].trim() !== '') {
                engineers.push(project[engineerKey]);
            } else {
                break;
            }
        }
        engineers = getEngineerOrder(project['项目'], engineers);
        // 轮流且集中分配值班任务，优先本项目内值班天数最少的工程师
        const n = dates.length;
        const m = engineers.length;
        let schedule = [];
        let avg = Math.floor(n / m);
        let remain = n % m;
        let localDays = {};
        engineers.forEach(e => localDays[e] = avg);
        for (let i = 0; i < remain; i++) {
            let minEngineer = engineers[0];
            engineers.forEach(e => {
                if (localDays[e] < localDays[minEngineer]) {
                    minEngineer = e;
                }
            });
            localDays[minEngineer]++;
        }
        let idx2 = 0;
        engineers.forEach(e => {
            for (let j = 0; j < localDays[e]; j++) {
                schedule.push(e);
                idx2++;
            }
        });
        // 新增：生成每日排班数据
        let scheduleArr = [];
        if (type === 'week') {
            // 按周
            let cur = new Date(startDate);
            let idx = 0;
            while (cur <= end) {
                let main = engineers[idx % engineers.length];
                let backups = engineers.filter(e => e !== main).join('、');
                for (let i = 0; i < 7; i++) {
                    let d = new Date(cur.getTime() + i * 24 * 60 * 60 * 1000);
                    if (d > end) break;
                    scheduleArr.push({date: d.toISOString().slice(0,10), main, backups, remark: ''});
                }
                cur = new Date(cur.getTime() + 7 * 24 * 60 * 60 * 1000);
                idx++;
            }
        } else {
            // 按月
            let cur = new Date(startDate);
            let idx = 0;
            while (cur <= end) {
                let main = engineers[idx % engineers.length];
                let backups = engineers.filter(e => e !== main).join('、');
                scheduleArr.push({date: cur.toISOString().slice(0,10), main, backups, remark: ''});
                cur.setMonth(cur.getMonth() + 1);
                idx++;
            }
        }
        dailyScheduleData[proj['项目']] = scheduleArr;
        // 横向表格生成
        const projTable = document.createElement('table');
        projTable.className = 'table table-bordered mb-4';
        // 表头
        let theadHtml = '<thead><tr><th>日期</th>';
        dates.forEach(date => {
            theadHtml += `<th>${date}</th>`;
        });
        theadHtml += '</tr></thead>';
        // 值班工程师行
        let ondutyHtml = '<tr><th>值班工程师</th>';
        schedule.forEach(eng => {
            ondutyHtml += `<td>${eng}</td>`;
        });
        ondutyHtml += '</tr>';
        // 备选工程师行
        let backupHtml = '<tr><th>备选工程师</th>';
        schedule.forEach(eng => {
            backupHtml += `<td>${engineers.filter(e => e !== eng).join('、') || '无'}</td>`;
        });
        backupHtml += '</tr>';
        projTable.innerHTML = theadHtml + '<tbody>' + ondutyHtml + backupHtml + '</tbody>';
        cardBody.appendChild(projTable);
        card.appendChild(cardBody);
        container.appendChild(card);
    });
    // 初始化折叠功能（Bootstrap 5）
    setTimeout(() => {
        const collapseEls = container.querySelectorAll('.collapse');
        collapseEls.forEach((el, i) => {
            if (i === 0) {
                // 默认展开第一个
                new bootstrap.Collapse(el, {toggle: true});
            } else {
                new bootstrap.Collapse(el, {toggle: false});
            }
        });
    }, 0);
}

// 监听项目筛选下拉框变化，重新生成排班表
function initProjectFilterListener() {
    const filter = document.getElementById('projectFilter');
    if (filter) {
        filter.addEventListener('change', generateSchedule);
    }
}

// 导出排班表
function exportSchedule() {
    // 获取所有项目卡片
    const cards = document.querySelectorAll('.card.mb-3');
    if (!cards || cards.length === 0) {
        showCustomAlert('请先生成排班表！', '提示');
        return;
    }
    // 创建新的工作簿
    const wb = XLSX.utils.book_new();
    cards.forEach((card, idx) => {
        // 获取项目名和项目信息
        let sheetName = '项目' + (idx + 1);
        let projectTitle = '';
        let project = null;
        // 获取卡片头部
        const cardHeader = card.querySelector('.card-header');
        if (cardHeader) {
            projectTitle = cardHeader.textContent.trim();
            const match = projectTitle.match(/项目：(.+?)（/);
            if (match) {
                sheetName = match[1];
                project = projectData.find(p => p['项目'] === match[1]);
            } else {
                sheetName = projectTitle.replace(/项目：/, '').trim();
                project = projectData.find(p => p['项目'] === sheetName);
            }
        }
        if (!project) {
            console.error('未找到项目数据:', sheetName);
            return;
        }
        // 获取卡片内容
        const cardBody = card.querySelector('.card-body, .collapse');
        if (!cardBody) return;
        // 获取表格
        const table = cardBody.querySelector('table');
        if (!table) return;
        // 解析表格内容为二维数组
        const aoa = [];
        const rows = table.querySelectorAll('tr');
        let colCount = 0;
        if (rows.length > 0) {
            colCount = rows[0].children.length;
        }
        // 插入项目信息，合并单元格
        aoa.push([projectTitle, ...Array(colCount - 1).fill('')]);
        // 将项目信息分成两行显示
        const smInfo = `SM电话号码：${project['SM电话号码'] || ''}    项目邮箱组：${project['项目邮箱组'] || ''}`;
        const callCenterInfo = `Call Center联系方式：${project['Call Center联系方式'] || ''}`;
        aoa.push([smInfo, ...Array(colCount - 1).fill('')]);
        aoa.push([callCenterInfo, ...Array(colCount - 1).fill('')]);
        rows.forEach(row => {
            const cells = Array.from(row.children).map(cell => cell.textContent.trim());
            aoa.push(cells);
        });
        // 生成sheet
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        // 合并前三行单元格
        ws['!merges'] = [
            { s: { r: 0, c: 0 }, e: { r: 0, c: colCount - 1 } },
            { s: { r: 1, c: 0 }, e: { r: 1, c: colCount - 1 } },
            { s: { r: 2, c: 0 }, e: { r: 2, c: colCount - 1 } }
        ];
        // 设置样式：微软雅黑10.5号，第一行加粗，前三行左对齐
        for (let R = 0; R < aoa.length; ++R) {
            for (let C = 0; C < colCount; ++C) {
                const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                if (!ws[cellRef]) continue;
                ws[cellRef].s = {
                    font: {
                        name: '微软雅黑',
                        sz: 10.5,
                        bold: R === 0 // 第一行加粗
                    },
                    alignment: {
                        vertical: 'center',
                        horizontal: (R <= 2) ? 'left' : 'center',
                        wrapText: true
                    }
                };
            }
        }
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, '排班表.xlsx');
}

// 上传项目信息
function handleFileUpload(event) {
    const file = event.target.files[0];
    if (file) {
        const reader = new FileReader();
        reader.onload = function(e) {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            let rawData = XLSX.utils.sheet_to_json(firstSheet, {defval: ''});
            // 保证每条数据都有"项目分值"字段
            rawData = rawData.map(row => {
                if (!('项目分值' in row)) row['项目分值'] = '';
                return row;
            });
            projectData = rawData;
            // 保存到localStorage
            localStorage.setItem('projectData', JSON.stringify(projectData));
            updateDashboard();
            if (currentPage === 'project-distribution') {
                updateProjectList();
            } else if (currentPage === 'personal-stats') {
                updatePersonStats();
            }
        };
        reader.readAsArrayBuffer(file);
    }
}

// 导出项目信息表
function exportProjectInfo() {
    if (!projectData || projectData.length === 0) {
        showCustomAlert('没有可导出的项目信息表！', '提示');
        return;
    }
    // 保证导出时"项目分值"在"项目"后、"服务经理"前
    const headers = [];
    if (projectData.length > 0) {
        const keys = Object.keys(projectData[0]);
        // 先找"项目"
        const idx = keys.indexOf('项目');
        if (idx !== -1) {
            headers.push(...keys.slice(0, idx + 1));
            headers.push('项目分值');
            keys.forEach(k => {
                if (k !== '项目' && k !== '项目分值') headers.push(k);
            });
        } else {
            headers.push(...keys);
        }
    }
    // 重新组织数据
    const exportData = projectData.map(row => {
        const newRow = {};
        headers.forEach(h => { newRow[h] = row[h] || ''; });
        return newRow;
    });
    const ws = XLSX.utils.json_to_sheet(exportData, {header: headers});
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, '运维项目信息表');
    XLSX.writeFile(wb, '运维项目信息表.xlsx');
}

// 修改初始化事件监听器，生成排班表前先更新筛选下拉框
function initEventListeners() {
    // 项目搜索
    document.getElementById('projectSearch').addEventListener('input', function(e) {
        const searchText = e.target.value.toLowerCase();
        const items = document.querySelectorAll('#projectList .list-group-item');
        items.forEach(item => {
            const text = item.textContent.toLowerCase();
            item.style.display = text.includes(searchText) ? '' : 'none';
        });
    });

    // 人员搜索
    document.getElementById('personSearch').addEventListener('input', updatePersonStats);
    document.getElementById('personType').addEventListener('change', updatePersonStats);

    // 生成排班表
    document.getElementById('generateSchedule').addEventListener('click', function() {
        updateProjectFilter();
        generateSchedule();
        generateEngineerSchedule();
        initProjectFilterListener();
    });

    // 导出排班表
    document.getElementById('exportSchedule').addEventListener('click', exportSchedule);

    // 文件上传
    document.getElementById('projectInfoFile').addEventListener('change', handleFileUpload);
    document.getElementById('uploadProjectInfo').addEventListener('click', function() {
        document.getElementById('projectInfoFile').click();
    });

    // 排班规则显示/隐藏
    const toggleBtn = document.getElementById('toggleRules');
    if (toggleBtn) {
        toggleBtn.addEventListener('click', function() {
            const rules = document.getElementById('scheduleRules');
            if (rules.style.display === 'none') {
                rules.style.display = 'block';
            } else {
                rules.style.display = 'none';
            }
        });
    }

    // 导出项目信息表
    document.getElementById('exportProjectInfo').addEventListener('click', exportProjectInfo);
}

// 按工程师维度展示排班
function generateEngineerSchedule() {
    const area = document.getElementById('engineerScheduleArea');
    area.innerHTML = '';
    // 获取用户选择的日期
    const startDateInput = document.getElementById('startDate').value;
    const endDateInput = document.getElementById('endDate').value;
    if (!startDateInput || !endDateInput) return;
    const startDate = new Date(startDateInput);
    const endDate = new Date(endDateInput);
    if (startDate > endDate) return;
    // 生成日期范围
    const dates = [];
    const currentDate = new Date(startDate);
    while (currentDate <= endDate) {
        dates.push(currentDate.toISOString().split('T')[0]);
        currentDate.setDate(currentDate.getDate() + 1);
    }
    // 统计所有工程师的排班
    const engineerMap = {};
    projectData.forEach(project => {
        // 获取本项目的排班计划
        const engineers = [];
        for (let i = 1; ; i++) {
            const engineerKey = `工程师${i}`;
            if (project[engineerKey] && project[engineerKey].trim() !== '') {
                engineers.push(project[engineerKey]);
            } else {
                break;
            }
        }
        const n = dates.length;
        const m = engineers.length;
        let schedule = [];
        let avg = Math.floor(n / m);
        let remain = n % m;
        let localDays = {};
        engineers.forEach(e => localDays[e] = avg);
        for (let i = 0; i < remain; i++) {
            let minEngineer = engineers[0];
            engineers.forEach(e => {
                if (localDays[e] < localDays[minEngineer]) {
                    minEngineer = e;
                }
            });
            localDays[minEngineer]++;
        }
        let idx = 0;
        engineers.forEach(e => {
            for (let j = 0; j < localDays[e]; j++) {
                schedule.push(e);
                idx++;
            }
        });
        // 记录每个工程师每天负责的项目
        dates.forEach((date, i) => {
            const eng = schedule[i];
            if (!engineerMap[eng]) engineerMap[eng] = {};
            if (!engineerMap[eng][date]) engineerMap[eng][date] = [];
            engineerMap[eng][date].push(project['项目']);
        });
    });
    // 工程师列表
    const allEngineers = Object.keys(engineerMap);
    const selectDiv = document.createElement('div');
    selectDiv.className = 'mb-3';
    // 原生多选checkbox模式
    let selectHtml = '<label class="me-2"><strong>选择工程师：</strong></label>';
    selectHtml += '<div id="engineerCheckboxList" style="display:inline-block;max-width:100%;overflow-x:auto;white-space:nowrap;vertical-align:middle;">';
    allEngineers.forEach(eng => {
        selectHtml += `<label class='me-2' style='white-space:nowrap;'><input type='checkbox' value='${eng}' class='form-check-input' style='margin-right:4px;'>${eng}</label>`;
    });
    selectHtml += '</div>';
    selectDiv.innerHTML = selectHtml;
    area.appendChild(selectDiv);
    // 展示选中的工程师排班
    function renderSelected() {
        // 移除旧表格
        area.querySelectorAll('.engineer-card').forEach(e => e.remove());
        const checked = Array.from(area.querySelectorAll('#engineerCheckboxList input:checked')).map(e => e.value);
        checked.forEach(eng => {
            const card = document.createElement('div');
            card.className = 'mb-3 p-2 border rounded engineer-card';
            // 折叠按钮
            const collapseId = 'engCollapse_' + eng.replace(/[^a-zA-Z0-9]/g, '');
            card.innerHTML = `<div class="d-flex align-items-center mb-2"><strong>${eng}</strong>
                <button class="btn btn-link btn-sm ms-2" data-bs-toggle="collapse" data-bs-target="#${collapseId}">展开/收起</button></div>
                <div class="collapse show" id="${collapseId}">
                    <table class="table table-bordered table-sm mt-2"><thead><tr><th>日期</th><th>负责项目</th></tr></thead><tbody>
                    ${dates.map(date => `<tr><td>${date}</td><td>${engineerMap[eng][date] ? engineerMap[eng][date].join('、') : ''}</td></tr>`).join('')}
                    </tbody></table>
                </div>`;
            area.appendChild(card);
        });
    }
    // 监听checkbox变化
    area.querySelector('#engineerCheckboxList').addEventListener('change', renderSelected);
    renderSelected();
}

function initDailySchedulePage() {
    // 初始化项目下拉
    const projectSelect = document.getElementById('dailyProjectSelect');
    if (projectSelect) {
        projectSelect.innerHTML = '<option value="all">全部项目</option>';
        const names = [...new Set(projectData.map(p => p['项目']))];
        names.forEach(name => {
            const opt = document.createElement('option');
            opt.value = name;
            opt.textContent = name;
            projectSelect.appendChild(opt);
        });
        projectSelect.onchange = function() {
            renderEngineerOrderArea();
        };
    }
    // 初始化时渲染一次
    renderEngineerOrderArea();
    // 监听生成按钮
    const genBtn = document.getElementById('generateDailySchedule');
    if (genBtn) {
        genBtn.onclick = function() {
            renderDailyScheduleTable();
        };
    }
    // 默认起始日期为今天
    const startDateInput = document.getElementById('dailyStartDate');
    if (startDateInput) {
        const today = new Date();
        startDateInput.value = today.toISOString().slice(0, 10);
    }
    // 日历视图按钮事件绑定
    const calendarBtn = document.getElementById('toggleCalendarView');
    if (calendarBtn) {
        calendarBtn.onclick = function() {
            const area = document.getElementById('calendarViewArea');
            const tableArea = document.getElementById('dailyScheduleTableArea');
            if (area.style.display === 'none' || area.style.display === '') {
                area.style.display = 'block';
                if (tableArea) tableArea.style.display = 'none';
                calendarBtn.classList.remove('btn-outline-primary');
                calendarBtn.classList.add('btn-primary', 'text-white');
                // 初次打开时，显示起始日期所在月
                const startDateInput = document.getElementById('dailyStartDate');
                let startDate = startDateInput && startDateInput.value ? startDateInput.value : (new Date()).toISOString().slice(0, 10);
                let start = new Date(startDate);
                renderCalendarView(start);
            } else {
                area.style.display = 'none';
                if (tableArea) tableArea.style.display = '';
                calendarBtn.classList.remove('btn-primary', 'text-white');
                calendarBtn.classList.add('btn-outline-primary');
            }
        };
    }
    // 日历视图切换月份按钮
    const prevBtn = document.getElementById('calendarPrevMonth');
    const nextBtn = document.getElementById('calendarNextMonth');
    if (prevBtn) {
        prevBtn.onclick = function() {
            if (!calendarViewMonth) return;
            let prev = new Date(calendarViewMonth.getFullYear(), calendarViewMonth.getMonth() - 1, 1);
            renderCalendarView(prev);
        };
    }
    if (nextBtn) {
        nextBtn.onclick = function() {
            if (!calendarViewMonth) return;
            let next = new Date(calendarViewMonth.getFullYear(), calendarViewMonth.getMonth() + 1, 1);
            renderCalendarView(next);
        };
    }
    // 监听排班类型选择，显示/隐藏自定义周期输入框
    const typeSelect = document.getElementById('dailyTypeSelect');
    const customCycleArea = document.getElementById('customCycleDaysArea');
    if (typeSelect && customCycleArea) {
        typeSelect.onchange = function() {
            if (typeSelect.value === 'custom') {
                customCycleArea.style.display = '';
            } else {
                customCycleArea.style.display = 'none';
            }
            renderEngineerOrderArea();
        };
    }
}

function getEngineerOrderKey(projectName) {
    return 'engineerOrder_' + encodeURIComponent(projectName || 'all');
}

function renderEngineerOrderArea() {
    const area = document.getElementById('engineerOrderArea');
    const project = document.getElementById('dailyProjectSelect').value;
    if (!project || project === 'all') {
        area.innerHTML = '<div class="text-muted">请选择具体项目以调整工程师顺序</div>';
        return;
    }
    // 获取工程师
    const proj = projectData.find(p => p['项目'] === project);
    if (!proj) {
        area.innerHTML = '<div class="text-danger">未找到该项目</div>';
        return;
    }
    let engineers = [];
    for (let i = 1; ; i++) {
        const key = `工程师${i}`;
        if (proj[key] && proj[key].trim() !== '') {
            engineers.push(proj[key]);
        } else {
            break;
        }
    }
    // 读取本地顺序
    const orderKey = getEngineerOrderKey(project);
    let savedOrder = localStorage.getItem(orderKey);
    if (savedOrder) {
        try {
            const arr = JSON.parse(savedOrder);
            if (Array.isArray(arr) && arr.length === engineers.length && arr.every(e => engineers.includes(e))) {
                engineers = arr;
            }
        } catch {}
    }
    // 渲染顺序列表
    let html = '<label class="form-label">工程师轮值顺序（可调整）：</label>';
    html += '<ul class="list-group" id="engineerOrderList">';
    engineers.forEach((eng, idx) => {
        html += `<li class="list-group-item d-flex align-items-center justify-content-between">
            <span>${idx + 1}. ${eng}</span>
            <span>
                <button type="button" class="btn btn-sm btn-outline-secondary me-1" data-action="up" data-idx="${idx}" ${idx === 0 ? 'disabled' : ''}>↑</button>
                <button type="button" class="btn btn-sm btn-outline-secondary" data-action="down" data-idx="${idx}" ${idx === engineers.length - 1 ? 'disabled' : ''}>↓</button>
            </span>
        </li>`;
    });
    html += '</ul>';
    area.innerHTML = html;
    // 绑定事件
    const list = area.querySelector('#engineerOrderList');
    list.addEventListener('click', function(e) {
        if (e.target.tagName === 'BUTTON') {
            const idx = parseInt(e.target.getAttribute('data-idx'));
            const action = e.target.getAttribute('data-action');
            if (action === 'up' && idx > 0) {
                [engineers[idx - 1], engineers[idx]] = [engineers[idx], engineers[idx - 1]];
            } else if (action === 'down' && idx < engineers.length - 1) {
                [engineers[idx + 1], engineers[idx]] = [engineers[idx], engineers[idx + 1]];
            }
            // 保存顺序
            localStorage.setItem(orderKey, JSON.stringify(engineers));
            renderEngineerOrderArea();
        }
    });
}

function getEngineerOrder(project, engineers) {
    if (!project || project === 'all') return engineers;
    const orderKey = getEngineerOrderKey(project);
    let savedOrder = localStorage.getItem(orderKey);
    if (savedOrder) {
        try {
            const arr = JSON.parse(savedOrder);
            if (Array.isArray(arr) && arr.length === engineers.length && arr.every(e => engineers.includes(e))) {
                return arr;
            }
        } catch {}
    }
    return engineers;
}

function renderDailyScheduleTable() {
    // 保存当前排班表为历史记录（带时间戳）
    // 新增：生成排班数据并保存到dailyScheduleData
    dailyScheduleData = {};
    const area = document.getElementById('dailyScheduleTableArea');
    const project = document.getElementById('dailyProjectSelect').value;
    const type = document.getElementById('dailyTypeSelect').value;
    // 起始日期由用户选择，排一年
    const startDateInput = document.getElementById('dailyStartDate');
    let startDate = startDateInput && startDateInput.value ? startDateInput.value : (new Date()).toISOString().slice(0, 10);
    const start = new Date(startDate);
    const end = new Date(start);
    end.setFullYear(start.getFullYear() + 1);
    end.setDate(end.getDate() - 1); // 满一年
    const endDate = end.toISOString().slice(0, 10);
    let showProjects = projectData;
    if (project !== 'all') {
        showProjects = projectData.filter(p => p['项目'] === project);
    }
    // 生成周期
    let periods = [];
    if (type === 'week') {
        // 按周
        let cur = new Date(startDate);
        const end = new Date(endDate);
        while (cur <= end) {
            const weekStart = new Date(cur);
            const weekEnd = new Date(cur);
            weekEnd.setDate(weekEnd.getDate() + 6);
            if (weekEnd > end) weekEnd.setTime(end.getTime());
            periods.push({start: new Date(weekStart), end: new Date(weekEnd)});
            cur.setDate(cur.getDate() + 7);
        }
    } else {
        // 按月
        let cur = new Date(startDate);
        const end = new Date(endDate);
        cur.setDate(1);
        while (cur <= end) {
            const monthStart = new Date(cur);
            const monthEnd = new Date(cur.getFullYear(), cur.getMonth() + 1, 0);
            if (monthEnd > end) monthEnd.setTime(end.getTime());
            periods.push({start: new Date(monthStart), end: new Date(monthEnd)});
            cur.setMonth(cur.getMonth() + 1);
        }
    }
    // 渲染表格
    let html = '';
    for (let p = 0; p < showProjects.length; p++) {
        const proj = showProjects[p];
        let engineers = [];
        for (let i = 1; ; i++) {
            const key = `工程师${i}`;
            if (proj[key] && proj[key].trim() !== '') {
                engineers.push(proj[key]);
            } else {
                break;
            }
        }
        engineers = getEngineerOrder(proj['项目'], engineers);
        if (engineers.length === 0) continue;
        // 生成每日排班数据（本地变量）
        let scheduleArr = [];
        if (type === 'week') {
            let cur = new Date(startDate);
            let idx = 0;
            while (cur <= end) {
                let main = engineers[idx % engineers.length];
                let backups = engineers.filter(e => e !== main).join('、');
                for (let i = 0; i < 7; i++) {
                    let d = new Date(cur.getTime() + i * 24 * 60 * 60 * 1000);
                    if (d > end) break;
                    scheduleArr.push({date: d.toISOString().slice(0,10), main, backups, remark: ''});
                }
                cur = new Date(cur.getTime() + 7 * 24 * 60 * 60 * 1000);
                idx++;
            }
        } else if (type === 'month') {
            let cur = new Date(startDate);
            let idx = 0;
            while (cur <= end) {
                let main = engineers[idx % engineers.length];
                let backups = engineers.filter(e => e !== main).join('、');
                scheduleArr.push({date: cur.toISOString().slice(0,10), main, backups, remark: ''});
                cur.setMonth(cur.getMonth() + 1);
                idx++;
            }
        } else if (type === 'custom') {
            const cycleInput = document.getElementById('customCycleDays');
            let cycleDays = parseInt(cycleInput && cycleInput.value ? cycleInput.value : '3', 10);
            if (isNaN(cycleDays) || cycleDays < 1) {
                showCustomAlert('轮班周期必须为大于等于1的整数！', '错误');
                continue;
            }
            let cur = new Date(startDate);
            let idx = 0;
            while (cur <= end) {
                let main = engineers[idx % engineers.length];
                let backups = engineers.filter(e => e !== main).join('、');
                for (let i = 0; i < cycleDays; i++) {
                    let d = new Date(cur.getTime() + i * 24 * 60 * 60 * 1000);
                    if (d > end) break;
                    scheduleArr.push({date: d.toISOString().slice(0,10), main, backups, remark: ''});
                }
                cur = new Date(cur.getTime() + cycleDays * 24 * 60 * 60 * 1000);
                idx++;
            }
        }
        // 渲染表格
        if (type === 'custom') {
            let rows = '';
            scheduleArr.forEach((item, idx) => {
                rows += `<tr data-project="${proj['项目']}" data-idx="${idx}">
                    <td>${item.date}</td>
                    <td class="main-engineer">${item.main}</td>
                    <td class="backup-engineer">${item.backups}</td>
                    <td class="remark">${item.remark || ''}</td>
                    <td><button class="btn btn-sm btn-link edit-row-btn">编辑</button></td>
                </tr>`;
            });
            html += `<div class="table-responsive mb-4"><table class="table table-bordered"><thead><tr><th>日期</th><th>值班工程师</th><th>备选工程师</th><th>备注</th><th>操作</th></tr></thead><tbody>${rows}</tbody></table></div>`;
            dailyScheduleData[proj['项目']] = scheduleArr;
            continue;
        }
        // ...原有渲染表格逻辑...
        dailyScheduleData[proj['项目']] = scheduleArr;
        let rows = '';
        periods.forEach((period, idx) => {
            let main = engineers[idx % engineers.length];
            let backups = engineers.filter(e => e !== main).join('、');
            let periodStr = type === 'week'
                ? `${period.start.getFullYear()}-${period.start.getMonth()+1}/${period.start.getDate()} ~ ${period.end.getFullYear()}-${period.end.getMonth()+1}/${period.end.getDate()}`
                : `${period.start.getFullYear()}-${(period.start.getMonth()+1).toString().padStart(2,'0')}`;
            rows += `<tr data-project="${proj['项目']}" data-idx="${idx}">
                <td>${proj['项目']}</td>
                <td>${periodStr}</td>
                <td class="main-engineer">${main}</td>
                <td class="backup-engineer">${backups}</td>
                <td class="remark"></td>
                <td><button class="btn btn-sm btn-link edit-row-btn">编辑</button></td>
            </tr>`;
        });
        html += `<div class="table-responsive mb-4"><table class="table table-bordered"><thead><tr><th>项目名称</th><th>${type==='week'?'周期（起止日期）':'月份'}</th><th>值班工程师</th><th>备选工程师</th><th>备注</th><th>操作</th></tr></thead><tbody>${rows}</tbody></table></div>`;
    }
    if (!html) html = '<div class="text-muted">无排班数据</div>';
    area.innerHTML = html;
    // 渲染完成后再保存历史记录
    saveCurrentDailyScheduleToHistory();
    // ...原有编辑按钮绑定逻辑...
    setTimeout(() => {
        const tables = area.querySelectorAll('table');
        tables.forEach(table => {
            table.querySelectorAll('.edit-row-btn').forEach(btn => {
                btn.onclick = function() {
                    const tr = btn.closest('tr');
                    if (!tr) return;
                    if (btn.textContent === '编辑') {
                        // ...原有编辑逻辑...
                    } else if (btn.textContent === '保存') {
                        // 保存编辑内容
                        const tr = btn.closest('tr');
                        const mainTd = tr.querySelector('.main-engineer');
                        const backupTd = tr.querySelector('.backup-engineer');
                        const remarkTd = tr.querySelector('.remark');
                        const mainVal = mainTd.querySelector('.main-edit').value;
                        const projectName = tr.getAttribute('data-project');
                        // 备选工程师自动生成
                        let proj = projectData.find(p => p['项目'] === projectName);
                        let engineers = [];
                        for (let i = 1; ; i++) {
                            const key = `工程师${i}`;
                            if (proj && proj[key] && proj[key].trim() !== '') {
                                engineers.push(proj[key]);
                            } else {
                                break;
                            }
                        }
                        engineers = getEngineerOrder(projectName, engineers);
                        const backupVals = engineers.filter(e => e !== mainVal);
                        const remarkVal = remarkTd.querySelector('.remark-edit').value;
                        mainTd.textContent = mainVal;
                        backupTd.textContent = backupVals.join('、');
                        remarkTd.textContent = remarkVal;
                        btn.textContent = '编辑';
                        // 移除撤销按钮
                        const cancelBtn = tr.querySelector('.cancel-edit-btn');
                        if (cancelBtn) cancelBtn.remove();
                        // 自动保存为历史记录
                        saveCurrentDailyScheduleToHistory();
                        // 新增：同步更新dailyScheduleData
                        // 通过周期idx找到对应的日期范围，更新dailyScheduleData
                        const idx = parseInt(tr.getAttribute('data-idx'));
                        const type = document.getElementById('dailyTypeSelect').value;
                        if (type === 'week') {
                            // 周排班，更新该周7天
                            let arr = dailyScheduleData[projectName];
                            if (arr) {
                                for (let i = 0; i < 7; i++) {
                                    let arrIdx = idx * 7 + i;
                                    if (arr[arrIdx]) {
                                        arr[arrIdx].main = mainVal;
                                        arr[arrIdx].backups = backupVals.join('、');
                                        arr[arrIdx].remark = remarkVal;
                                    }
                                }
                            }
                        } else {
                            // 月排班，更新该月
                            let arr = dailyScheduleData[projectName];
                            if (arr && arr[idx]) {
                                arr[idx].main = mainVal;
                                arr[idx].backups = backupVals.join('、');
                                arr[idx].remark = remarkVal;
                            }
                        }
                    }
                };
            });
        });
    }, 0);
}

function saveCurrentDailyScheduleToHistory() {
    // 获取当前排班表HTML和参数
    const area = document.getElementById('dailyScheduleTableArea');
    if (!area) return;
    const html = area.innerHTML;
    if (!html || html.includes('无排班数据')) return;
    // 保存参数（项目、类型、起始日期、顺序等）
    const project = document.getElementById('dailyProjectSelect').value;
    const type = document.getElementById('dailyTypeSelect').value;
    const startDateInput = document.getElementById('dailyStartDate');
    let startDate = startDateInput && startDateInput.value ? startDateInput.value : (new Date()).toISOString().slice(0, 10);
    // 组装历史记录对象
    const record = {
        time: new Date().toISOString(),
        project,
        type,
        startDate,
        html
    };
    // 存入localStorage
    let history = [];
    try {
        history = JSON.parse(localStorage.getItem('dailyScheduleHistory') || '[]');
    } catch {}
    history.push(record);
    localStorage.setItem('dailyScheduleHistory', JSON.stringify(history));
}

function exportDailySchedule() {
    // 获取当前排班表数据
    const area = document.getElementById('dailyScheduleTableArea');
    const tables = area.querySelectorAll('table');
    if (!tables || tables.length === 0) {
        showCustomAlert('请先生成排班表！', '提示');
        return;
    }
    // 多项目导出到一个Excel文件，每个项目一个sheet
    const wb = XLSX.utils.book_new();
    tables.forEach((table, idx) => {
        // 获取项目名
        let sheetName = '项目' + (idx + 1);
        const caption = table.querySelector('tr td');
        if (caption) {
            sheetName = caption.textContent.trim();
        }
        // 解析表格内容为二维数组
        const aoa = [];
        const rows = table.querySelectorAll('tr');
        rows.forEach(row => {
            const cells = Array.from(row.children).map(cell => cell.textContent.trim());
            aoa.push(cells);
        });
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
    });
    XLSX.writeFile(wb, '各项目日常排班表.xlsx');
}

function renderDailyHistoryList() {
    const listDiv = document.getElementById('dailyHistoryList');
    let history = [];
    try {
        history = JSON.parse(localStorage.getItem('dailyScheduleHistory') || '[]');
    } catch {}
    if (!history.length) {
        listDiv.innerHTML = '<div class="text-muted">暂无历史记录</div>';
        return;
    }
    // 倒序显示，最新在前
    history = history.slice().reverse();
    let html = '<div class="list-group">';
    history.forEach((item, idx) => {
        const time = new Date(item.time).toLocaleString();
        html += `<a href="#" class="list-group-item list-group-item-action" data-idx="${history.length-1-idx}">
            <div class="d-flex w-100 justify-content-between">
                <h6 class="mb-1">${item.project === 'all' ? '全部项目' : item.project} | ${item.type === 'week' ? '周排班' : '月排班'} | 起始：${item.startDate}</h6>
                <small>${time}</small>
            </div>
            <div class="small text-muted">点击查看详情</div>
        </a>`;
    });
    html += '</div>';
    listDiv.innerHTML = html;
    // 绑定点击事件，弹窗内可查看详情
    listDiv.querySelectorAll('.list-group-item').forEach(a => {
        a.onclick = function(e) {
            e.preventDefault();
            const idx = parseInt(this.getAttribute('data-idx'));
            showDailyHistoryDetail(idx);
        };
    });
}

function showDailyHistoryDetail(idx) {
    let history = [];
    try {
        history = JSON.parse(localStorage.getItem('dailyScheduleHistory') || '[]');
    } catch {}
    if (!history[idx]) return;
    const item = history[idx];
    const listDiv = document.getElementById('dailyHistoryList');
    let html = `<div class="mb-2"><strong>项目：</strong>${item.project === 'all' ? '全部项目' : item.project} &nbsp; <strong>类型：</strong>${item.type === 'week' ? '周排班' : '月排班'} &nbsp; <strong>起始：</strong>${item.startDate}<br><strong>生成时间：</strong>${new Date(item.time).toLocaleString()}</div>`;
    html += '<div class="mb-2">' + item.html + '</div>';
    html += '<button class="btn btn-sm btn-secondary me-2" id="backToHistoryList">返回列表</button>';
    html += '<button class="btn btn-sm btn-primary" id="restoreHistoryRecord">恢复为当前排班表</button>';
    listDiv.innerHTML = html;
    document.getElementById('backToHistoryList').onclick = renderDailyHistoryList;
    document.getElementById('restoreHistoryRecord').onclick = function() {
        restoreDailyHistoryRecord(idx);
    };
}

function restoreDailyHistoryRecord(idx) {
    let history = [];
    try {
        history = JSON.parse(localStorage.getItem('dailyScheduleHistory') || '[]');
    } catch {}
    if (!history[idx]) return;
    // 恢复html到当前排班表区域
    document.getElementById('dailyScheduleTableArea').innerHTML = history[idx].html;
    // 关闭模态框
    const modal = bootstrap.Modal.getInstance(document.getElementById('dailyHistoryModal'));
    if (modal) modal.hide();
}

// 通用自定义提示框
function showCustomAlert(message, title = '提示') {
    document.getElementById('customAlertTitle').textContent = title;
    document.getElementById('customAlertBody').textContent = message;
    const modal = new bootstrap.Modal(document.getElementById('customAlertModal'));
    modal.show();
}

// 日历视图渲染
function renderCalendarView(monthDate) {
    const area = document.getElementById('calendarViewArea');
    const tableArea = document.getElementById('calendarTableArea') || area;
    tableArea.innerHTML = '';
    // 获取当前筛选的项目、类型、起始日期
    const project = document.getElementById('dailyProjectSelect').value;
    const type = document.getElementById('dailyTypeSelect').value;
    const startDateInput = document.getElementById('dailyStartDate');
    let startDate = startDateInput && startDateInput.value ? startDateInput.value : (new Date()).toISOString().slice(0, 10);
    let showProjects = projectData;
    if (project !== 'all') {
        showProjects = projectData.filter(p => p['项目'] === project);
    }
    // 取当前要渲染的月
    let start = new Date(startDate);
    let year, month;
    if (monthDate) {
        year = monthDate.getFullYear();
        month = monthDate.getMonth();
    } else {
        year = start.getFullYear();
        month = start.getMonth();
    }
    calendarViewMonth = new Date(year, month, 1);
    const label = document.getElementById('calendarMonthLabel');
    if (label) label.textContent = `${year}年${(month+1).toString().padStart(2,'0')}月`;
    const firstDay = new Date(year, month, 1);
    const lastDay = new Date(year, month + 1, 0);
    let html = '<div class="table-responsive"><table class="table table-bordered text-center"><thead><tr>';
    const weekDays = ['日','一','二','三','四','五','六'];
    weekDays.forEach(d => html += `<th>${d}</th>`);
    html += '</tr></thead><tbody>';
    showProjects.forEach(proj => {
        // 直接用dailyScheduleData
        let arr = dailyScheduleData[proj['项目']] || [];
        // 过滤出本月数据
        let days = arr.filter(item => {
            const d = new Date(item.date);
            return d.getFullYear() === year && d.getMonth() === month;
        });
        // 渲染日历
        html += `<tr><td colspan="7" class="bg-light text-start"><strong>${proj['项目']}</strong></td></tr>`;
        let cur = 0;
        for (let w = 0; w < 6; w++) { // 最多6周
            html += '<tr>';
            for (let wd = 0; wd < 7; wd++) {
                if (w === 0 && wd < firstDay.getDay()) {
                    html += '<td></td>';
                } else if (cur < days.length) {
                    if (days[cur].main) {
                        html += `<td style="min-width:90px;">
                            <div><strong>${new Date(days[cur].date).getDate()}</strong></div>
                            <div style="font-size:0.9em;">${days[cur].main}</div>
                            <div style="font-size:0.8em;color:#888;">${days[cur].backups}</div>
                        </td>`;
                    } else {
                        html += `<td style="min-width:90px;"><div><strong>${new Date(days[cur].date).getDate()}</strong></div></td>`;
                    }
                    cur++;
                } else {
                    html += '<td></td>';
                }
            }
            html += '</tr>';
            if (cur >= days.length) break;
        }
    });
    html += '</tbody></table></div>';
    tableArea.innerHTML = html;
}

// 导出为按天表格
if (document.getElementById('exportDailyByDay')) {
    document.getElementById('exportDailyByDay').onclick = function() {
        // 获取当前筛选的项目、类型、起始日期
        const project = document.getElementById('dailyProjectSelect').value;
        const type = document.getElementById('dailyTypeSelect').value;
        let showProjects = projectData;
        if (project !== 'all') {
            showProjects = projectData.filter(p => p['项目'] === project);
        }
        // 直接用dailyScheduleData
        let aoa = [['项目名称','日期','值班工程师','备选工程师','备注']];
        showProjects.forEach(proj => {
            let arr = dailyScheduleData[proj['项目']] || [];
            arr.forEach(item => {
                aoa.push([
                    proj['项目'],
                    item.date,
                    item.main,
                    item.backups,
                    item.remark || ''
                ]);
            });
        });
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, '按天排班表');
        XLSX.writeFile(wb, '按天排班表.xlsx');
    };
}

// 导出为月视图表格
if (document.getElementById('exportDailyMonthView')) {
    document.getElementById('exportDailyMonthView').onclick = function() {
        // 获取当前筛选的项目、类型、起始日期
        const project = document.getElementById('dailyProjectSelect').value;
        const type = document.getElementById('dailyTypeSelect').value;
        let showProjects = projectData;
        if (project !== 'all') {
            showProjects = projectData.filter(p => p['项目'] === project);
        }
        const weekDays = ['日','一','二','三','四','五','六'];
        const wb = XLSX.utils.book_new();
        showProjects.forEach(proj => {
            let arr = dailyScheduleData[proj['项目']] || [];
            // 按月分组
            let monthMap = {};
            arr.forEach(item => {
                const d = new Date(item.date);
                const key = `${d.getFullYear()}-${(d.getMonth()+1).toString().padStart(2,'0')}`;
                if (!monthMap[key]) monthMap[key] = [];
                monthMap[key].push(item);
            });
            Object.keys(monthMap).forEach(monthKey => {
                let days = monthMap[monthKey];
                let year = parseInt(monthKey.split('-')[0]);
                let month = parseInt(monthKey.split('-')[1]) - 1;
                let firstDay = new Date(year, month, 1);
                let lastDay = new Date(year, month + 1, 0);
                let aoa = [];
                aoa.push([monthKey]);
                aoa.push(weekDays);
                let cur = 0;
                for (let w = 0; w < 6; w++) {
                    let row = [];
                    for (let wd = 0; wd < 7; wd++) {
                        if (w === 0 && wd < firstDay.getDay()) {
                            row.push('');
                        } else if (cur < days.length) {
                            if (days[cur].main) {
                                row.push(`${new Date(days[cur].date).getDate()}
${days[cur].main}`);
                            } else {
                                row.push(`${new Date(days[cur].date).getDate()}`);
                            }
                            cur++;
                        } else {
                            row.push('');
                        }
                    }
                    aoa.push(row);
                    if (cur >= days.length) break;
                }
                const ws = XLSX.utils.aoa_to_sheet(aoa);
                // 设置所有单元格自动换行
                const range = XLSX.utils.decode_range(ws['!ref']);
                for (let R = 0; R <= range.e.r; ++R) {
                    for (let C = 0; C <= range.e.c; ++C) {
                        const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                        if (ws[cellRef]) {
                            if (!ws[cellRef].s) ws[cellRef].s = {};
                            if (!ws[cellRef].s.alignment) ws[cellRef].s.alignment = {};
                            ws[cellRef].s.alignment.wrapText = true;
                        }
                    }
                }
                XLSX.utils.book_append_sheet(wb, ws, monthKey);
            });
        });
        XLSX.writeFile(wb, '月视图排班表.xlsx');
    };
} 