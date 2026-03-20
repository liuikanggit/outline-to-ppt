/**
 * 🕹️ Outline to PPT - Frontend Logic (PyWebView Bridge)
 */

document.addEventListener('DOMContentLoaded', () => {
    const generateBtn = document.getElementById('generateBtn');
    const shareIdInput = document.getElementById('shareId');
    const consolePanel = document.getElementById('consolePanel');
    const page1 = document.getElementById('page1');
    const page2 = document.getElementById('page2');
    
    const openFolderBtn = document.getElementById('openFolderBtn');
    const openPptBtn = document.getElementById('openPptBtn');
    const backBtn = document.getElementById('backBtn');

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
            alert("执行失败：" + message);
        }
        resetBtn();
    };

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

    // 文件夹/PPT 按钮绑定
    openFolderBtn.addEventListener('click', () => {
        const shareId = shareIdInput.value.trim();
        if (window.pywebview && window.pywebview.api) {
            window.pywebview.api.open_output_folder(shareId);
        }
    });

    openPptBtn.addEventListener('click', () => {
        const shareId = shareIdInput.value.trim();
        if (window.pywebview && window.pywebview.api) {
            window.pywebview.api.open_final_ppt(shareId);
        }
    });


});

