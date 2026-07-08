document.addEventListener("DOMContentLoaded", () => {
    // UI 元素
    const excelPathInput = document.getElementById("excel-path");
    const btnBrowseExcel = document.getElementById("btn-browse-excel");
    const dragArea = document.getElementById("drag-area");
    const fileUploader = document.getElementById("file-uploader");
    const targetUrlInput = document.getElementById("target-url");
    
    const startRowInput = document.getElementById("start-row");
    const maxRowInput = document.getElementById("max-row");
    
    const colNameInput = document.getElementById("col-name");
    const colExamIdInput = document.getElementById("col-exam-id");
    const colPasswordInput = document.getElementById("col-password");
    const colScoreInput = document.getElementById("col-score");
    
    const extractModeSelect = document.getElementById("extract-mode");
    const extractParamInput = document.getElementById("extract-param");
    
    const chkOverwrite = document.getElementById("chk-overwrite");
    const chkShowBrowser = document.getElementById("chk-show-browser");
    const chkExportNew = document.getElementById("chk-export-new");
    const exportPathGroup = document.getElementById("export-path-group");
    const exportPathInput = document.getElementById("export-path");
    const btnBrowseExport = document.getElementById("btn-browse-export");

    // 高级定位选择器
    const selExamId = document.getElementById("sel-exam-id");
    const selPassword = document.getElementById("sel-password");
    const selCaptchaInput = document.getElementById("sel-captcha-input");
    const selCaptchaImg = document.getElementById("sel-captcha-img");
    const selQueryBtn = document.getElementById("sel-query-btn");

    // 运行按钮与控制台
    const themeToggle = document.getElementById("theme-toggle");
    const btnStart = document.getElementById("btn-start");
    const btnStop = document.getElementById("btn-stop");
    const statusBadge = document.getElementById("status-badge");
    const previewRowCount = document.getElementById("preview-row-count");
    const previewTableBody = document.querySelector("#preview-table tbody");
    const consoleBox = document.getElementById("console-box");

    // 状态状态标志
    let isRunning = false;
    let pollIntervalId = null;
    let lastLogIndex = 0;

    // 1. 初始化加载
    const currentTheme = localStorage.getItem("theme");
    if (currentTheme === "dark") {
        document.body.classList.add("dark-mode");
    }

    themeToggle.addEventListener("click", () => {
        document.body.classList.toggle("dark-mode");
        if (document.body.classList.contains("dark-mode")) {
            localStorage.setItem("theme", "dark");
        } else {
            localStorage.setItem("theme", "light");
        }
    });

    loadConfig();
    
    // 2. 事件监听

    // 提取模式变更
    extractModeSelect.addEventListener("change", (e) => {
        const mode = e.target.value;
        extractParamInput.disabled = false;
        if (mode === "自动智能匹配") {
            extractParamInput.value = "无参数";
            extractParamInput.disabled = true;
        } else if (mode.includes("行号切片范围")) {
            extractParamInput.value = "62:78";
        } else if (mode.includes("自定义")) {
            extractParamInput.value = "table td";
        }
    });

    // 导出文件切换
    chkExportNew.addEventListener("change", () => {
        if (chkExportNew.checked) {
            exportPathGroup.style.display = "block";
        } else {
            exportPathGroup.style.display = "none";
        }
    });

    // 本地浏览文件
    btnBrowseExcel.addEventListener("click", () => {
        fetch("/api/browse/excel")
            .then(res => res.json())
            .then(data => {
                if (data.path) {
                    excelPathInput.value = data.path;
                    loadPreview();
                }
            });
    });

    btnBrowseExport.addEventListener("click", () => {
        fetch("/api/browse/folder")
            .then(res => res.json())
            .then(data => {
                if (data.path) {
                    exportPathInput.value = data.path;
                }
            });
    });

    // 当路径输入框失焦时自动刷新表格预览
    excelPathInput.addEventListener("change", loadPreview);

    // 拖拽 Excel 上传
    dragArea.addEventListener("click", () => fileUploader.click());
    
    dragArea.addEventListener("dragover", (e) => {
        e.preventDefault();
        dragArea.classList.add("dragover");
    });

    dragArea.addEventListener("dragleave", () => {
        dragArea.classList.remove("dragover");
    });

    dragArea.addEventListener("drop", (e) => {
        e.preventDefault();
        dragArea.classList.remove("dragover");
        const files = e.dataTransfer.files;
        if (files.length > 0 && files[0].name.endsWith(".xlsx")) {
            uploadExcel(files[0]);
        }
    });

    fileUploader.addEventListener("change", () => {
        if (fileUploader.files.length > 0) {
            uploadExcel(fileUploader.files[0]);
        }
    });

    // 按钮控制：启动查询
    btnStart.addEventListener("click", () => {
        if (!excelPathInput.value) {
            alert("请先选择或拖拽上传 Excel 数据表！");
            return;
        }
        
        const config = getFormConfig();
        
        // 禁用控件
        setControlsEnabled(false);
        isRunning = true;
        
        // 清空控制台
        consoleBox.innerHTML = '<div class="console-line text-success">系统启动中，正在创建查分线程...</div>';
        lastLogIndex = 0;

        fetch("/api/start", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(config)
        })
        .then(res => res.json())
        .then(data => {
            if (data.status === "started") {
                btnStop.disabled = false;
                // 开始高频轮询状态
                startPolling();
            } else {
                alert("启动失败: " + data.message);
                setControlsEnabled(true);
            }
        })
        .catch(err => {
            alert("请求失败: " + err);
            setControlsEnabled(true);
        });
    });

    // 按钮控制：停止查询
    btnStop.addEventListener("click", () => {
        btnStop.disabled = true;
        appendLogLine("正在发送停止查询请求，请等待当前学生处理完...", "warn");
        
        fetch("/api/stop", { method: "POST" })
            .then(res => res.json())
            .then(data => {
                appendLogLine("停止请求发送成功: " + data.message, "success");
            });
    });

    // 3. 通用辅助函数

    // 运行期间禁用/启用配置项
    function setControlsEnabled(enabled) {
        btnStart.disabled = !enabled;
        excelPathInput.disabled = !enabled;
        btnBrowseExcel.disabled = !enabled;
        targetUrlInput.disabled = !enabled;
        startRowInput.disabled = !enabled;
        maxRowInput.disabled = !enabled;
        colNameInput.disabled = !enabled;
        colExamIdInput.disabled = !enabled;
        colPasswordInput.disabled = !enabled;
        colScoreInput.disabled = !enabled;
        extractModeSelect.disabled = !enabled;
        extractParamInput.disabled = !enabled;
        chkOverwrite.disabled = !enabled;
        chkShowBrowser.disabled = !enabled;
        chkExportNew.disabled = !enabled;
        exportPathInput.disabled = !enabled;
        btnBrowseExport.disabled = !enabled;
        dragArea.style.pointerEvents = enabled ? "auto" : "none";
        dragArea.style.opacity = enabled ? "1" : "0.5";
    }

    // 从页面收集配置
    function getFormConfig() {
        return {
            excel_path: excelPathInput.value,
            target_url: targetUrlInput.value,
            column_mapping: {
                name: parseInt(colNameInput.value),
                exam_id: parseInt(colExamIdInput.value),
                password: parseInt(colPasswordInput.value),
                first_score: parseInt(colScoreInput.value)
            },
            start_row: parseInt(startRowInput.value),
            max_row: parseInt(maxRowInput.value),
            overwrite: chkOverwrite.checked,
            export_new: chkExportNew.checked,
            export_path: exportPathInput.value,
            headless: !chkShowBrowser.checked, // 选中为显示，因此 headless 取反
            selectors: {
                exam_id: selExamId.value,
                password: selPassword.value,
                captcha_input: selCaptchaInput.value,
                captcha_img: selCaptchaImg.value,
                query_btn: selQueryBtn.value
            },
            extraction_mode: extractModeSelect.value,
            extraction_param: extractParamInput.value
        };
    }

    // 后端上传 Excel
    function uploadExcel(file) {
        const formData = new FormData();
        formData.append("file", file);

        appendLogLine("正在上传 Excel 文件: " + file.name, "success");
        
        fetch("/api/upload", {
            method: "POST",
            body: formData
        })
        .then(res => res.json())
        .then(data => {
            if (data.status === "success") {
                excelPathInput.value = data.path;
                appendLogLine("Excel 上传并解析成功！", "success");
                loadPreview();
            } else {
                alert("上传失败: " + data.message);
            }
        })
        .catch(err => alert("上传发生网络错误: " + err));
    }

    // 读取后端 Excel 配置
    function loadConfig() {
        fetch("/api/config")
            .then(res => res.json())
            .then(data => {
                if (data.excel_path) {
                    excelPathInput.value = data.excel_path;
                }
                if (data.target_url) {
                    targetUrlInput.value = data.target_url;
                }
                if (data.column_mapping) {
                    colNameInput.value = data.column_mapping.name || 2;
                    colExamIdInput.value = data.column_mapping.exam_id || 3;
                    colPasswordInput.value = data.column_mapping.password || 5;
                    colScoreInput.value = data.column_mapping.first_score || 6;
                }
                if (data.start_row) startRowInput.value = data.start_row;
                if (data.max_row) maxRowInput.value = data.max_row;
                
                chkOverwrite.checked = data.overwrite || false;
                chkShowBrowser.checked = !data.headless; // headless 为 false 则是显示浏览器
                chkExportNew.checked = data.export_new || false;
                
                if (chkExportNew.checked) {
                    exportPathGroup.style.display = "block";
                }
                if (data.export_path) {
                    exportPathInput.value = data.export_path;
                }

                // 高级选择器
                if (data.selectors) {
                    if (data.selectors.exam_id) selExamId.value = data.selectors.exam_id;
                    if (data.selectors.password) selPassword.value = data.selectors.password;
                    if (data.selectors.captcha_input) selCaptchaInput.value = data.selectors.captcha_input;
                    if (data.selectors.captcha_img) selCaptchaImg.value = data.selectors.captcha_img;
                    if (data.selectors.query_btn) selQueryBtn.value = data.selectors.query_btn;
                }

                // 提取器
                if (data.extraction_mode) {
                    extractModeSelect.value = data.extraction_mode;
                    extractParamInput.value = data.extraction_param || "无参数";
                    if (data.extraction_mode === "自动智能匹配") {
                        extractParamInput.disabled = true;
                    } else {
                        extractParamInput.disabled = false;
                    }
                }

                // 初始化完配置后载入数据表格预览
                loadPreview();
            });
    }

    // 加载学生名单表格预览
    function loadPreview() {
        const path = excelPathInput.value;
        if (!path) return;

        previewRowCount.innerText = "读取中...";

        fetch(`/api/excel/preview?path=${encodeURIComponent(path)}&name_col=${colNameInput.value}&id_col=${colExamIdInput.value}&pwd_col=${colPasswordInput.value}&score_col=${colScoreInput.value}`)
            .then(res => res.json())
            .then(data => {
                if (data.status === "success") {
                    renderTable(data.students);
                    previewRowCount.innerText = `共读取 ${data.students.length} 行考生记录`;
                } else {
                    previewTableBody.innerHTML = `<tr><td colspan="6" class="text-center text-error">加载预览失败: ${data.message}</td></tr>`;
                    previewRowCount.innerText = "读取失败";
                }
            })
            .catch(err => {
                previewTableBody.innerHTML = `<tr><td colspan="6" class="text-center text-error">加载错误: ${err}</td></tr>`;
            });
    }

    // 渲染预览表格
    function renderTable(students) {
        if (!students || students.length === 0) {
            previewTableBody.innerHTML = `<tr><td colspan="6" class="text-center text-muted">表格中没有读取到有效的考生记录</td></tr>`;
            return;
        }

        let html = "";
        students.forEach(s => {
            let statusText = "等待查询";
            let statusClass = "status-pending";
            
            if (s.status === "running") {
                statusText = "查询中";
                statusClass = "status-running";
            } else if (s.status === "success") {
                statusText = "已完成";
                statusClass = "status-success";
            } else if (s.status === "error") {
                statusText = "查询失败";
                statusClass = "status-error";
            } else if (s.status === "skipped") {
                statusText = "已跳过";
                statusClass = "status-skipped";
            }

            let scoresHtml = "";
            if (s.scores && s.scores.length > 0) {
                s.scores.forEach(val => {
                    scoresHtml += `<span class="score-badge">${val}</span>`;
                });
            } else if (s.message) {
                scoresHtml = `<span class="text-muted" style="font-size: 0.8rem;">${s.message}</span>`;
            } else {
                scoresHtml = '<span class="text-muted">-</span>';
            }

            html += `
                <tr id="row-${s.row}">
                    <td>${s.row}</td>
                    <td><strong>${s.name || "-"}</strong></td>
                    <td class="text-muted">${s.exam_id || "-"}</td>
                    <td class="text-muted">${s.password || "-"}</td>
                    <td><span class="status-indicator ${statusClass}">${statusText}</span></td>
                    <td>${scoresHtml}</td>
                </tr>
            `;
        });
        previewTableBody.innerHTML = html;
    }

    // 4. 运行状态轮询

    function startPolling() {
        if (pollIntervalId) clearInterval(pollIntervalId);
        
        statusBadge.className = "badge badge-running";
        statusBadge.innerText = "正在查询中...";

        pollIntervalId = setInterval(() => {
            fetch(`/api/status?last_log=${lastLogIndex}`)
                .then(res => res.json())
                .then(data => {
                    // 更新每个学生的状态行
                    if (data.students && data.students.length > 0) {
                        data.students.forEach(s => {
                            const tr = document.getElementById(`row-${s.row}`);
                            if (tr) {
                                // 更新状态单元格
                                const statusTd = tr.querySelector("td:nth-child(5)");
                                let statusText = "等待查询";
                                let statusClass = "status-pending";
                                
                                if (s.status === "running") {
                                    statusText = "查询中";
                                    statusClass = "status-running";
                                    // 自动滚动表格使正在查询的行处于可见区
                                    tr.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                                } else if (s.status === "success") {
                                    statusText = "已完成";
                                    statusClass = "status-success";
                                } else if (s.status === "error") {
                                    statusText = "查询失败";
                                    statusClass = "status-error";
                                } else if (s.status === "skipped") {
                                    statusText = "已跳过";
                                    statusClass = "status-skipped";
                                }
                                statusTd.innerHTML = `<span class="status-indicator ${statusClass}">${statusText}</span>`;

                                // 更新分数单元格
                                const scoreTd = tr.querySelector("td:nth-child(6)");
                                let scoresHtml = "";
                                if (s.scores && s.scores.length > 0) {
                                    s.scores.forEach(val => {
                                        scoresHtml += `<span class="score-badge">${val}</span>`;
                                    });
                                } else if (s.message) {
                                    scoresHtml = `<span class="text-muted" style="font-size: 0.8rem;">${s.message}</span>`;
                                } else {
                                    scoresHtml = '<span class="text-muted">-</span>';
                                }
                                scoreTd.innerHTML = scoresHtml;
                            }
                        });
                    }

                    // 写入新增的日志行
                    if (data.logs && data.logs.length > 0) {
                        data.logs.forEach(log => {
                            appendLogLine(log.text, log.level);
                        });
                        lastLogIndex = data.last_index;
                    }

                    // 查询完毕处理
                    if (!data.is_running) {
                        clearInterval(pollIntervalId);
                        pollIntervalId = null;
                        isRunning = false;
                        setControlsEnabled(true);
                        btnStop.disabled = true;
                        
                        statusBadge.className = "badge badge-idle";
                        statusBadge.innerText = "系统空闲";
                        appendLogLine("查询进程已全部结束，资源已释放。", "success");
                    }
                })
                .catch(err => {
                    console.error("轮询状态接口失败:", err);
                });
        }, 1000);
    }

    function appendLogLine(text, level) {
        const line = document.createElement("div");
        line.className = "console-line";
        if (level === "error" || level === "critical") {
            line.classList.add("text-error");
        } else if (level === "warning" || level === "warn") {
            line.classList.add("text-warn");
        } else if (level === "success" || text.includes("完成") || text.includes("成功")) {
            line.classList.add("text-success");
        }
        line.innerText = text;
        consoleBox.appendChild(line);
        consoleBox.scrollTop = consoleBox.scrollHeight;
    }
});
