// 员工离职数据分析 - 前端脚本
var currentMonth = null;
var chartDept, chartType, chartTenure, chartReason, chartLevel;

function showLoading() { document.getElementById('loadingOverlay').style.display = 'flex'; }
function hideLoading() { document.getElementById('loadingOverlay').style.display = 'none'; }

document.getElementById('uploadZone').addEventListener('click', function() { document.getElementById('fileInput').click(); });
document.getElementById('uploadZone').addEventListener('dragover', function(e) { e.preventDefault(); this.style.borderColor = '#2a5298'; });
document.getElementById('uploadZone').addEventListener('dragleave', function() { this.style.borderColor = '#ccc'; });
document.getElementById('uploadZone').addEventListener('drop', function(e) { e.preventDefault(); this.style.borderColor = '#ccc'; if (e.dataTransfer.files[0]) handleFileUpload(e.dataTransfer.files[0]); });
document.getElementById('fileInput').addEventListener('change', function() { if (this.files[0]) handleFileUpload(this.files[0]); });

function handleFileUpload(file) {
    if (!file) return;
    showLoading();
    var formData = new FormData();
    formData.append('file', file);
    fetch('/api/upload', { method: 'POST', body: formData })
        .then(function(r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
        .then(function(data) {
            hideLoading();
            if (data.success) {
                document.getElementById('uploadStatus').innerHTML = '<p style="color:green;"><i class="bi bi-check-circle"></i> ' + data.message + '</p><p>检测到月份: ' + data.summary.available_months.join(', ') + '</p>';
                loadMonths(data.summary.month_options);
            } else {
                document.getElementById('uploadStatus').innerHTML = '<p style="color:red;"><i class="bi bi-x-circle"></i> ' + data.message + '</p>';
            }
        })
        .catch(function(err) { hideLoading(); document.getElementById('uploadStatus').innerHTML = '<p style="color:red;"><i class="bi bi-x-circle"></i> 上传失败: ' + err.message + '</p>'; });
}

function loadMonths(options) {
    var select = document.getElementById('monthSelect');
    select.innerHTML = '';
    options.forEach(function(opt) {
        var optEl = document.createElement('option');
        optEl.value = opt.value;
        optEl.text = opt.label;
        select.appendChild(optEl);
    });
    document.getElementById('filterSection').classList.remove('hidden');
    document.getElementById('statsRow').classList.remove('hidden');
    document.getElementById('chartsRow1').classList.remove('hidden');
    document.getElementById('chartsRow2').classList.remove('hidden');
    document.getElementById('chartsRow3').classList.remove('hidden');
    if (options.length > 0) loadMonthData();
}

function loadMonthData() {
    currentMonth = document.getElementById('monthSelect').value;
    if (!currentMonth) return;
    showLoading();
    fetch('/api/analyze', { method: 'POST', headers: {'Content-Type': 'application/json'}, body: JSON.stringify({month: currentMonth}) })
        .then(function(r) { if (!r.ok) throw new Error('HTTP ' + r.status); return r.json(); })
        .then(function(data) {
            hideLoading();
            if (data.success) { updateUI(data.data); } else { alert('加载数据失败: ' + data.message); }
        })
        .catch(function(err) { hideLoading(); alert('加载失败: ' + err.message); });
}

function updateUI(data) {
    document.getElementById('statMonth').textContent = data.month || '-';
    document.getElementById('statAvgCount').textContent = data.avg_count || '-';
    document.getElementById('statTurnoverCount').textContent = data.turnover_count || '-';
    document.getElementById('statTurnoverRate').textContent = (data.turnover_rate || '0') + '%';
    renderCharts(data);
}

function renderCharts(data) {
    // 部门离职率柱状图+折线图
    if (chartDept) chartDept.dispose();
    chartDept = echarts.init(document.getElementById('chartDept'));
    var deptData = data.dept_turnover || [];
    chartDept.setOption({
        tooltip: { trigger: 'axis', formatter: function(params) { var d = params[0]; var r = params[1] ? params[1].value : 0; return d.name + '<br/>离职人数: ' + d.value + '<br/>离职率: ' + r + '%'; } },
        legend: { data: ['离职人数', '离职率'] },
        xAxis: { type: 'category', data: deptData.map(function(d) { return d['一级组织']; }), axisLabel: { rotate: 30, interval: 0 } },
        yAxis: [{ type: 'value', name: '人数', position: 'left' }, { type: 'value', name: '离职率(%)', position: 'right', max: 100, axisLabel: { formatter: '{value}%' } }],
        series: [
            { name: '离职人数', type: 'bar', data: deptData.map(function(d) { return d['离职人数']; }), itemStyle: { color: '#2a5298' }, label: { show: true, position: 'top' } },
            { name: '离职率', type: 'line', yAxisIndex: 1, data: deptData.map(function(d) { return d['离职率'] || 0; }), itemStyle: { color: '#ff6b6b' }, lineStyle: { width: 2 }, symbol: 'circle', label: { show: true, formatter: '{c}%' } }
        ]
    });

    // 离职类型饼图
    if (chartType) chartType.dispose();
    chartType = echarts.init(document.getElementById('chartType'));
    var typeData = data.turnover_type || [];
    chartType.setOption({
        tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
        series: [{ type: 'pie', radius: ['40%', '70%'], data: typeData.map(function(d) { return { name: d['离职类型'], value: d['人数'] }; }), label: { formatter: '{b}: {d}%' } }]
    });

    // 司龄分布饼图
    if (chartTenure) chartTenure.dispose();
    chartTenure = echarts.init(document.getElementById('chartTenure'));
    var tenureData = data.turnover_tenure || [];
    chartTenure.setOption({
        tooltip: { trigger: 'item', formatter: '{b}: {c} ({d}%)' },
        series: [{ type: 'pie', radius: ['40%', '70%'], data: tenureData.map(function(d) { return { name: d['司龄段'], value: d['人数'] }; }), label: { formatter: '{b}: {d}%' } }]
    });

    // 离职原因柱状图
    if (chartReason) chartReason.dispose();
    chartReason = echarts.init(document.getElementById('chartReason'));
    var reasonData = data.turnover_reason || [];
    chartReason.setOption({
        tooltip: { trigger: 'axis' },
        xAxis: { type: 'category', data: reasonData.map(function(d) { return d['离职原因']; }), axisLabel: { rotate: 30 } },
        yAxis: { type: 'value', name: '人数' },
        series: [{ type: 'bar', data: reasonData.map(function(d) { return d['人数']; }), itemStyle: { color: '#1e3c72' }, label: { show: true, position: 'top' } }]
    });

    // 职级分布柱状图
    if (chartLevel) chartLevel.dispose();
    chartLevel = echarts.init(document.getElementById('chartLevel'));
    var levelData = data.turnover_level || [];
    chartLevel.setOption({
        tooltip: { trigger: 'axis' },
        xAxis: { type: 'category', data: levelData.map(function(d) { return d['职级']; }), axisLabel: { rotate: 30 } },
        yAxis: { type: 'value', name: '人数' },
        series: [{ type: 'bar', data: levelData.map(function(d) { return d['人数']; }), itemStyle: { color: '#28a745' }, label: { show: true, position: 'top' } }]
    });
}

function exportExcel() {
    if (!currentMonth) { alert('请先选择月份'); return; }

function exportExcel() {
    if (!currentMonth) { alert("请先选择月份"); return; }
    fetch("/api/export", { method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify({month: currentMonth}) })
        .then(function(r) { return r.blob(); })
        .then(function(blob) {
            var url = window.URL.createObjectURL(blob);
            var a = document.createElement("a");
            a.href = url;
            a.download = currentMonth + "离职分析.xlsx";
            a.click();
            window.URL.revokeObjectURL(url);
        })
        .catch(function(err) { alert("导出失败: " + err.message); });
}

function clearData() { location.reload(); }

window.addEventListener("resize", function() {
    if (chartDept) chartDept.resize();
    if (chartType) chartType.resize();
    if (chartTenure) chartTenure.resize();
    if (chartReason) chartReason.resize();
    if (chartLevel) chartLevel.resize();
});

