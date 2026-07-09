/**
 * 🕹️ Outline to PPT - Frontend Logic (PyWebView Bridge)
 */

document.addEventListener('DOMContentLoaded', () => {
    const generateBtn = document.getElementById('generateBtn');
    const shareIdInput = document.getElementById('shareId');
    const consolePanel = document.getElementById('consolePanel');
    const page1 = document.getElementById('page1');
    const page2 = document.getElementById('page2');
    
    const saveAsBtn = document.getElementById('saveAsBtn');
    const openPptBtn = document.getElementById('openPptBtn');
    const backBtn = document.getElementById('backBtn');

    const page3 = document.getElementById('page3');
    const errorStep = document.getElementById('errorStep');
    const errorCategory = document.getElementById('errorCategory');
    const errorLocationRow = document.getElementById('errorLocationRow');
    const errorLocation = document.getElementById('errorLocation');
    const errorAdvice = document.getElementById('errorAdvice');
    const errorSummary = document.getElementById('errorSummary');
    const openLogBtn = document.getElementById('openLogBtn');
    const backFromErrorBtn = document.getElementById('backFromErrorBtn');

    // 状态管理
    let isProcessing = false;

    // 绑定点击事件
    generateBtn.addEventListener('click', async () => {
        const shareId = shareIdInput.value.trim();
        if (!shareId) {
            alert('请输入 ShareID 或 链接！');
            return;
        }

        if (isProcessing) return;
        isProcessing = true;

        // UI 锁死和展现第一页状态
        generateBtn.disabled = true;
        generateBtn.style.opacity = '0.6';
        generateBtn.innerText = '⏳ 正在组装幻灯片...';
        consolePanel.style.display = 'block';

        const consoleLog = document.getElementById('consoleLog');
        if (consoleLog) consoleLog.innerText = '任务已启动，请耐心等待...';

        // 调用 Python API
        try {
            if (window.pywebview && window.pywebview.api) {
                const res = await window.pywebview.api.generate_ppt(shareId);
                // 这里API其实只是触发了后台进程
                if (res.status !== 'success') {
                    alert('生成启动失败: ' + (res.message || '未知错误'));
                    resetBtn();
                }
            } else {
                alert('PyWebView API 未加载，请在客户端中运行！');
                resetBtn();
            }
        } catch (err) {
            console.error(err);
            alert('调用出错: ' + err.message);
            resetBtn();
        }
    });

    // 恢复按钮状态
    function resetBtn() {
        isProcessing = false;
        generateBtn.disabled = false;
        generateBtn.style.opacity = '1';
        generateBtn.innerHTML = '<span class="btn-text">🚀 一键生成 PPT</span><span class="btn-glow"></span>';
    }

    // 🖥️ Python 调用的回调函数 (当 Python 任务全部完结时调用)
    window.taskCompleted = (isSuccess, message) => {
        if (isSuccess) {
            // 切页
            page1.style.display = 'none';
            page2.style.display = 'block';
        } else {
            showErrorPage(message);
        }
        resetBtn();
    };

    // ❌ 展示结构化的失败信息，帮助用户判断是幕布内容问题还是需要联系开发者
    function showErrorPage(message) {
        let info = null;
        try {
            info = JSON.parse(message);
        } catch (e) {
            info = null;
        }

        if (info) {
            errorStep.innerText = "失败于：" + info.step;
            errorCategory.innerText = info.category;
            errorAdvice.innerText = info.advice || '';
            errorSummary.innerText = info.error_summary || '';
            openLogBtn.style.display = info.log_path ? 'block' : 'none';

            if (info.location) {
                errorLocation.innerText = info.location;
                errorLocationRow.style.display = 'block';
            } else {
                errorLocationRow.style.display = 'none';
            }
        } else {
            // 兜底：万一 Python 端传回的不是 JSON，退化为展示原始文本
            errorStep.innerText = '';
            errorCategory.innerText = '未知错误';
            errorLocationRow.style.display = 'none';
            errorAdvice.innerText = '请打开日志文件或联系开发者协助排查。';
            errorSummary.innerText = String(message);
            openLogBtn.style.display = 'block';
        }

        page1.style.display = 'none';
        page2.style.display = 'none';
        page3.style.display = 'block';
    }

    // 📄 打开报错日志文件所在位置
    openLogBtn.addEventListener('click', () => {
        if (window.pywebview && window.pywebview.api) {
            window.pywebview.api.open_error_log();
        }
    });

    // 返回重试
    backFromErrorBtn.addEventListener('click', () => {
        page3.style.display = 'none';
        page1.style.display = 'block';
        consolePanel.style.display = 'none';
    });

    // 📜 实时日志追加
    window.appendLog = (text) => {
        const consoleLog = document.getElementById('consoleLog');
        if (!consoleLog) return;
        
        // 如果是首次写入，清空默认文字
        if (consoleLog.innerText === '等待任务开始...' || consoleLog.innerText === '任务已启动，请耐心等待...') {
            consoleLog.innerText = '';
        }
        
        consoleLog.innerText += text + '\n';
        // 自动滚动到底部
        consoleLog.scrollTop = consoleLog.scrollHeight;
    };

    // 底部返回按钮
    backBtn.addEventListener('click', () => {
        page2.style.display = 'none';
        page1.style.display = 'block';
        consolePanel.style.display = 'none';
    });

    // 💾 另存为按钮导出
    saveAsBtn.addEventListener('click', () => {
        if (window.pywebview && window.pywebview.api) {
            window.pywebview.api.save_ppt_as();
        }
    });

    // 🖥️ 预览按钮
    openPptBtn.addEventListener('click', () => {
        if (window.pywebview && window.pywebview.api) {
            window.pywebview.api.open_final_ppt();
        }
    });

    // 从 Python 端获取统一版本号并渲染
    window.addEventListener('pywebviewready', function() {
        if (window.pywebview && window.pywebview.api) {
            window.pywebview.api.get_version().then(function(ver) {
                document.title = "Outline to PPT - 智能幻灯片生成器 v" + ver;
                const verSpan = document.getElementById('app-version');
                if (verSpan) verSpan.innerText = "v" + ver;
            });
        }
    });

});

