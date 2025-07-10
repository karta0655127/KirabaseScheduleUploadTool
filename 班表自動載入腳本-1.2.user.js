// ==UserScript==
// @name         班表自動載入腳本
// @namespace    https://www.instagram.com/yuxuan_0122_/
// @version      1.2
// @description  在右側新增滑鼠懸浮下拉視窗，含可切換是否自動隱藏的滑動開關，動畫轉換場景更加流暢
// @match        *://prod-kb-gm.kb.marscatgames.com.tw/*
// @grant        none
// @exclude      *://grok.com/*
// @exclude      *://chatgpt.com/*
// ==/UserScript==

(function() {
    'use strict';

    // 檢查是否為頂層視窗，若不是則終止執行
    if (window.top !== window.self) {
        return;
    }

    // 初始化時恢復數據
    let Ximen = {};
    let ShiftTimes = {};


    // 動態載入 jQuery（若頁面未包含）
    let jQueryLoaded = !!window.jQuery;
    if (!jQueryLoaded) {
        const script = document.createElement('script');
        script.src = 'https://code.jquery.com/jquery-3.7.1.min.js';
        script.async = true;
        script.onload = () => {
            jQueryLoaded = true;
        };
        script.onerror = () => {
            jQueryLoaded = false;
            console.error('無法載入 jQuery');
        };
        document.head.appendChild(script);
    }

    // 動態載入 SheetJS 庫
    let sheetJSLoaded = !!window.XLSX;
    if (!sheetJSLoaded) {
        const script = document.createElement('script');
        script.src = 'https://unpkg.com/xlsx@latest/dist/xlsx.full.min.js';
        script.async = true;
        script.onload = () => {
            sheetJSLoaded = true;
        };
        script.onerror = () => {
            sheetJSLoaded = false;
            document.getElementById('result-display').textContent = '無法載入 SheetJS 庫，請檢查網路、禁用廣告攔截器或重新整理頁面';
        };
        document.head.appendChild(script);
    }

    // 創建懸浮視窗的容器
    const floatingWindow = document.createElement('div');
    floatingWindow.id = 'floating-window';
    floatingWindow.style.position = 'fixed';
    floatingWindow.style.top = '0';
    floatingWindow.style.left = '50%';
    floatingWindow.style.transform = 'translateX(-50%)';
    floatingWindow.style.backgroundColor = '#f0f0f0';
    floatingWindow.style.border = '1px solid #ccc';
    floatingWindow.style.borderRadius = '5px';
    floatingWindow.style.padding = '10px';
    floatingWindow.style.zIndex = '10000';
    floatingWindow.style.width = '150px';
    floatingWindow.style.transition = 'all 0.5s ease'; // 絲滑高度轉換
    floatingWindow.style.boxShadow = '0 2px 5px rgba(0,0,0,0.2)';
    floatingWindow.style.fontFamily = 'Arial, sans-serif';
    floatingWindow.style.fontSize = '15px';

    // 標題
    const titleContainer = document.createElement('div'); // 使用容器來控制佈局
    titleContainer.style.display = 'flex';
    titleContainer.style.justifyContent = 'center';
    titleContainer.style.alignItems = 'center';

    const title = document.createElement('span');
    title.textContent = '班表載入工具';
    title.style.margin = '0';
    title.style.padding = '5px 0';
    title.style.color = '#000';
    title.style.fontFamily = 'Arial, sans-serif';
    title.style.fontSize = '16px';

    const version = document.createElement('span');
    version.textContent = 'v1.1';
    version.style.fontSize = '12px';
    version.style.color = '#666';

    titleContainer.appendChild(title);
    titleContainer.appendChild(version);
    floatingWindow.appendChild(titleContainer);

    // 自動隱藏開關容器
    const toggleContainer = document.createElement('div');
    toggleContainer.style.display = 'flex';
    toggleContainer.style.alignItems = 'center';
    toggleContainer.style.margin = '10px 0';
    toggleContainer.style.justifyContent = 'center';

    // 自動隱藏標籤
    const toggleLabel = document.createElement('span');
    toggleLabel.textContent = '自動隱藏';
    toggleLabel.style.color = '#000';
    toggleLabel.style.fontFamily = 'Arial, sans-serif';
    toggleLabel.style.fontSize = '14px';
    toggleContainer.appendChild(toggleLabel);

    // 滑動開關
    const toggleSwitch = document.createElement('div');
    toggleSwitch.style.width = '70px';
    toggleSwitch.style.height = '24px';
    toggleSwitch.style.backgroundColor = '#ccc';
    toggleSwitch.style.borderRadius = '12px';
    toggleSwitch.style.position = 'relative';
    toggleSwitch.style.cursor = 'pointer';
    toggleSwitch.style.marginLeft = '30px';
    toggleSwitch.style.transition = 'background-color 0.3s';

    const toggleSlider = document.createElement('div');
    toggleSlider.style.width = '20px';
    toggleSlider.style.height = '20px';
    toggleSlider.style.backgroundColor = '#fff';
    toggleSlider.style.borderRadius = '50%';
    toggleSlider.style.position = 'absolute';
    toggleSlider.style.top = '2px';
    toggleSlider.style.left = '2px';
    toggleSlider.style.transition = 'left 0.3s';

    const toggleText = document.createElement('span');
    toggleText.textContent = 'OFF';
    toggleText.style.position = 'absolute';
    toggleText.style.top = '50%';
    toggleText.style.right = '5px';
    toggleText.style.transform = 'translateY(-50%)';
    toggleText.style.color = '#000';
    toggleText.style.fontFamily = 'Arial, sans-serif';
    toggleText.style.fontSize = '14px';


    toggleSwitch.appendChild(toggleSlider);
    toggleSwitch.appendChild(toggleText);
    toggleContainer.appendChild(toggleSwitch);

    // 檔案載入容器
    const fileContainer = document.createElement('div');
    fileContainer.style.display = 'flex';
    fileContainer.style.alignItems = 'center';
    fileContainer.style.margin = '10px 0';
    fileContainer.style.justifyContent = 'center';

    // 載入班表標籤
    const fileLabel = document.createElement('span');
    fileLabel.textContent = '載入班表';
    fileLabel.style.color = '#000';
    fileLabel.style.fontFamily = 'Arial, sans-serif';
    fileLabel.style.fontSize = '14px';
    fileContainer.appendChild(fileLabel);

    // 檔案輸入按鈕
    const fileInput = document.createElement('input');
    fileInput.type = 'file';
    fileInput.accept = '.xlsx,.xls';
    fileInput.style.display = 'none'; // 隱藏原生檔案輸入

    // 自訂檔案載入按鈕
    const fileButton = document.createElement('button');
    fileButton.textContent = '選擇檔案';
    fileButton.style.width = '70px';
    fileButton.style.height = '24px';
    fileButton.style.marginLeft = '30px';
    fileButton.style.backgroundColor = '#fff';
    fileButton.style.border = '1px solid #ccc';
    fileButton.style.borderRadius = '4px';
    fileButton.style.cursor = 'pointer';
    fileButton.style.fontFamily = 'Arial, sans-serif';
    fileButton.style.fontSize = '14px';
    fileButton.style.color = '#000';
    fileButton.style.textAlign = 'center';
    fileButton.style.lineHeight = '22px';
    fileButton.style.boxSizing = 'border-box';
    fileButton.style.transition = 'all 0.2s ease';
    fileButton.style.outline = 'none';
    fileButton.onmouseover = function() {
        this.style.backgroundColor = '#e0e0e0';
        this.style.boxShadow = '0 1px 3px rgba(0,0,0,0.1)';
    };
    fileButton.onmouseout = function() {
        this.style.backgroundColor = '#fff';
        this.style.boxShadow = 'none';
    };
    fileButton.onmousedown = function() {
        this.style.backgroundColor = '#d0d0d0';
        this.style.transform = 'scale(0.95)';
    };
    fileButton.onmouseup = function() {
        this.style.backgroundColor = '#e0e0e0';
        this.style.transform = 'scale(1)';
    };

    fileContainer.appendChild(fileInput);
    fileContainer.appendChild(fileButton);

    // 檔案名稱顯示
    const fileNameDisplay = document.createElement('div');
    fileNameDisplay.textContent = '尚未選擇檔案';
    fileNameDisplay.style.color = '#000';
    fileNameDisplay.style.marginTop = '5px';
    fileNameDisplay.style.fontSize = '14px';
    fileNameDisplay.style.fontFamily = 'Arial, sans-serif';
    fileNameDisplay.style.textAlign = 'center';

    // 日期控制容器
    const dateControlContainer = document.createElement('div');
    dateControlContainer.style.display = 'flex';
    dateControlContainer.style.alignItems = 'center';
    dateControlContainer.style.margin = '10px 0';
    dateControlContainer.style.justifyContent = 'center';

    // 上一天按鈕
    const prevDayButton = document.createElement('button');
    prevDayButton.textContent = '上一天';
    prevDayButton.disabled = true;
    prevDayButton.style.width = '50px';
    prevDayButton.style.height = '24px';
    prevDayButton.style.backgroundColor = '#fff';
    prevDayButton.style.border = '1px solid #ccc';
    prevDayButton.style.borderRadius = '4px';
    prevDayButton.style.cursor = 'not-allowed';
    prevDayButton.style.fontFamily = 'Arial, sans-serif';
    prevDayButton.style.fontSize = '12px';
    prevDayButton.style.color = '#000';
    prevDayButton.style.textAlign = 'center';
    prevDayButton.style.lineHeight = '22px';
    prevDayButton.style.opacity = '0.5';
    prevDayButton.style.transition = 'all 0.2s ease';
    prevDayButton.style.outline = 'none';

    // 日期標籤
    const dateLabel = document.createElement('span');
    dateLabel.textContent = '日期';
    dateLabel.style.color = '#000';
    dateLabel.style.fontFamily = 'Arial, sans-serif';
    dateLabel.style.fontSize = '12px';
    dateLabel.style.margin = '0 5px';
    dateLabel.style.marginLeft = '10px';

    // 日期下拉選單
    const dateSelect = document.createElement('select');
    dateSelect.disabled = true;
    dateSelect.style.width = '50px';
    dateSelect.style.height = '24px';
    dateSelect.style.border = '1px solid #ccc';
    dateSelect.style.borderRadius = '4px';
    dateSelect.style.fontFamily = 'Arial, sans-serif';
    dateSelect.style.fontSize = '14px';
    dateSelect.style.cursor = 'not-allowed';
    dateSelect.style.opacity = '0.5';
    dateSelect.style.padding = '0 5px';

    // 下一天按鈕
    const nextDayButton = document.createElement('button');
    nextDayButton.textContent = '下一天';
    nextDayButton.disabled = true;
    nextDayButton.style.width = '50px';
    nextDayButton.style.height = '24px';
    nextDayButton.style.backgroundColor = '#fff';
    nextDayButton.style.border = '1px solid #ccc';
    nextDayButton.style.borderRadius = '4px';
    nextDayButton.style.cursor = 'not-allowed';
    nextDayButton.style.fontFamily = 'Arial, sans-serif';
    nextDayButton.style.fontSize = '12px';
    nextDayButton.style.color = '#000';
    nextDayButton.style.textAlign = 'center';
    nextDayButton.style.lineHeight = '22px';
    nextDayButton.style.opacity = '0.5';
    nextDayButton.style.transition = 'all 0.2s ease';
    nextDayButton.style.outline = 'none';
    nextDayButton.style.marginLeft = '10px';

    // 將日期控制元件添加到容器
    dateControlContainer.appendChild(prevDayButton);
    dateControlContainer.appendChild(dateLabel);
    dateControlContainer.appendChild(dateSelect);
    dateControlContainer.appendChild(nextDayButton);

    // 店家控制容器
    const storeControlContainer = document.createElement('div');
    storeControlContainer.style.display = 'flex';
    storeControlContainer.style.alignItems = 'center';
    storeControlContainer.style.margin = '10px 0';
    storeControlContainer.style.justifyContent = 'center';

    // 選擇店家標籤
    const storeLabel = document.createElement('span');
    storeLabel.textContent = '選擇店家';
    storeLabel.style.color = '#000';
    storeLabel.style.fontFamily = 'Arial, sans-serif';
    storeLabel.style.fontSize = '12px';
    storeLabel.style.margin = '0 5px';
    storeLabel.style.marginLeft = '10px';

    // 店家下拉選單
    const storeSelect = document.createElement('select');
    storeSelect.disabled = false;
    storeSelect.style.width = '150px';
    storeSelect.style.height = '24px';
    storeSelect.style.border = '1px solid #ccc';
    storeSelect.style.borderRadius = '4px';
    storeSelect.style.fontFamily = 'Arial, sans-serif';
    storeSelect.style.fontSize = '12px';
    storeSelect.style.cursor = 'pointer';
    storeSelect.style.opacity = '1';
    storeSelect.style.padding = '0 5px';
    const storeOptions = ['台北西門基地', '台北三創基地', '台北信義基地', '台北車站基地', '幽靈水晶', 'All Team'];
    storeOptions.forEach(optionText => {
        const option = document.createElement('option');
        option.value = optionText;
        option.textContent = optionText;
        storeSelect.appendChild(option);
    });
    storeSelect.value = '台北西門基地'; // 預設選擇 All Team

    // 將店家控制元件添加到容器
    storeControlContainer.appendChild(storeLabel);
    storeControlContainer.appendChild(storeSelect);

    // 動作按鈕容器
    const actionButtonContainer = document.createElement('div');
    actionButtonContainer.style.display = 'flex';
    actionButtonContainer.style.alignItems = 'center';
    actionButtonContainer.style.margin = '10px 0';
    actionButtonContainer.style.justifyContent = 'center';

    // 輸入當日班表按鈕
    const inputDayButton = document.createElement('button');
    inputDayButton.textContent = '輸入當日班表';
    inputDayButton.disabled = true;
    inputDayButton.style.width = '100px';
    inputDayButton.style.height = '24px';
    inputDayButton.style.backgroundColor = '#fff';
    inputDayButton.style.border = '1px solid #ccc';
    inputDayButton.style.borderRadius = '4px';
    inputDayButton.style.cursor = 'not-allowed';
    inputDayButton.style.fontFamily = 'Arial, sans-serif';
    inputDayButton.style.fontSize = '12px';
    inputDayButton.style.color = '#000';
    inputDayButton.style.textAlign = 'center';
    inputDayButton.style.lineHeight = '22px';
    inputDayButton.style.opacity = '0.5';
    inputDayButton.style.transition = 'all 0.2s ease';
    inputDayButton.style.outline = 'none';

    // 清空當日班表按鈕
    const clearDayButton = document.createElement('button');
    clearDayButton.textContent = '清空當日班表';
    clearDayButton.disabled = false;
    clearDayButton.style.width = '100px';
    clearDayButton.style.height = '24px';
    clearDayButton.style.backgroundColor = '#fff';
    clearDayButton.style.border = '1px solid #ccc';
    clearDayButton.style.borderRadius = '4px';
    clearDayButton.style.cursor = 'pointer';
    clearDayButton.style.fontFamily = 'Arial, sans-serif';
    clearDayButton.style.fontSize = '12px';
    clearDayButton.style.color = '#000';
    clearDayButton.style.textAlign = 'center';
    clearDayButton.style.lineHeight = '22px';
    clearDayButton.style.opacity = '1';
    clearDayButton.style.transition = 'all 0.2s ease';
    clearDayButton.style.outline = 'none';
    clearDayButton.style.marginLeft = '10px';
    clearDayButton.onmouseover = function() {
        this.style.backgroundColor = '#e0e0e0';
        this.style.boxShadow = '0 1px 3px rgba(0,0,0,0.1)';
    };
    clearDayButton.onmouseout = function() {
        this.style.backgroundColor = '#fff';
        this.style.boxShadow = 'none';
    };
    clearDayButton.onmousedown = function() {
        this.style.backgroundColor = '#d0d0d0';
        this.style.transform = 'scale(0.95)';
    };
    clearDayButton.onmouseup = function() {
        this.style.backgroundColor = '#e0e0e0';
        this.style.transform = 'scale(1)';
    };

    // 將動作按鈕添加到容器
    actionButtonContainer.appendChild(inputDayButton);
    actionButtonContainer.appendChild(clearDayButton);

    // 解析結果顯示
    const resultDisplay = document.createElement('div');
    resultDisplay.style.color = '#000';
    resultDisplay.style.marginTop = '5px';
    resultDisplay.style.fontSize = '12px';
    resultDisplay.style.fontFamily = 'Arial, sans-serif';
    resultDisplay.style.textAlign = 'center';
    resultDisplay.style.maxHeight = '100px';
    resultDisplay.style.overflowY = 'auto';

    // 添加 by星辰王 文字
    const byText = document.createElement('div');
    byText.textContent = 'by星辰王';
    byText.style.position = 'absolute';
    byText.style.bottom = '2px';
    byText.style.left = '5px';
    byText.style.fontSize = '10px';
    byText.style.color = '#666';
    byText.style.cursor = 'pointer'; // 顯示手型游標
    byText.style.transition = 'color 0.3s, text-shadow 0.3s'; // 添加過渡效果
    byText.onmouseover = function() {
        byText.style.color = '#000';
        byText.style.textShadow = '0 0 5px rgba(0, 0, 0, 0.3)';
    };
    byText.onmouseout = function() {
        byText.style.color = '#666';
        byText.style.textShadow = 'none';
    };
    byText.onclick = function() {
        window.open('https://www.instagram.com/yuxuan_0122_/', '_blank'); // 點擊後打開指定網站
    };
    floatingWindow.appendChild(byText);

    // 將元素添加到視窗
    floatingWindow.appendChild(titleContainer);
    floatingWindow.appendChild(toggleContainer);
    floatingWindow.appendChild(fileContainer);
    floatingWindow.appendChild(fileNameDisplay);
    floatingWindow.appendChild(storeControlContainer);
    floatingWindow.appendChild(dateControlContainer);
    floatingWindow.appendChild(actionButtonContainer);
    floatingWindow.appendChild(resultDisplay);

    // 添加到頁面
    document.body.appendChild(floatingWindow);

    // 自動隱藏狀態
    let animationInProgress = false;
    let isAutoHideEnabled = false;
    let isExpanded = false;
    let hoverTimer = null;

    isExpanded = false;
    floatingWindow.style.width = '150px';
    toggleContainer.style.display = 'none';
    fileContainer.style.display = 'none';
    fileNameDisplay.style.display = 'none';
    dateControlContainer.style.display = 'none';
    storeControlContainer.style.display = 'none';
    actionButtonContainer.style.display = 'none';
    resultDisplay.style.display = 'none';
    byText.style.display = 'none';

    const containers = [
        toggleContainer,
        fileContainer,
        fileNameDisplay,
        dateControlContainer,
        storeControlContainer,
        actionButtonContainer,
        byText
    ];

    containers.forEach(container => {
        container.style.opacity = '0';
        container.style.transform = 'translateY(-10px)';
        container.style.transition = 'opacity 0.1s ease, transform 0.1s ease';
        container.style.display = 'none';
    });

    // 在初始化後自動觸發放大（模擬滑鼠移入）
    setTimeout(expandWindow, 1000); // 500ms 延遲放大，模擬自然展開

    // 開關點擊事件
    toggleSwitch.addEventListener('click', () => {
        isAutoHideEnabled = !isAutoHideEnabled;
        if (isAutoHideEnabled) {
            toggleSwitch.style.backgroundColor = '#4caf50';
            toggleSlider.style.left = '48px';
            toggleText.textContent = 'ON';
            toggleText.style.right = 'auto';
            toggleText.style.left = '10px';
        } else {
            toggleSwitch.style.backgroundColor = '#ccc';
            toggleSlider.style.left = '2px';
            toggleText.textContent = 'OFF';
            toggleText.style.left = 'auto';
            toggleText.style.right = '10px';
        }
    });

    // 自動放大邏輯
    /*function expandWindow() {
        if (!isExpanded) {
            isExpanded = true;
            floatingWindow.style.width = '300px';
            toggleContainer.style.display = 'flex';
            fileContainer.style.display = 'flex';
            fileNameDisplay.style.display = 'block';
            dateControlContainer.style.display = 'flex';
            storeControlContainer.style.display = 'flex';
            actionButtonContainer.style.display = 'flex';
            resultDisplay.style.display = 'none';
            byText.style.display = 'block';
        }
    }*/

    function expandWindow() {
        if (animationInProgress) return;
        animationInProgress = true;

        if (isExpanded) return;
        isExpanded = true;

        floatingWindow.style.width = '300px';
        resultDisplay.style.display = 'none';

        // 等待 0.3 秒再展開內容
        setTimeout(() => {
            const containers = [
                toggleContainer,
                fileContainer,
                fileNameDisplay,
                dateControlContainer,
                storeControlContainer,
                actionButtonContainer,
                byText
            ];

            containers.forEach((container, index) => {
                setTimeout(() => {
                    container.style.display = (container === fileNameDisplay || container === byText)
                        ? 'block' : 'flex';
                    requestAnimationFrame(() => {
                        container.style.opacity = '1';
                        container.style.transform = 'translateY(0)';
                    });
                    if (index === containers.length - 1) {
                        setTimeout(() => {
                            animationInProgress = false;
                        }, 100);
                    }
                }, index * 50); // 每個元素延遲 50ms
            });
        }, 300);
    }

    // 滑鼠移入展開
    floatingWindow.addEventListener('mouseenter', () => {
        if (hoverTimer) clearTimeout(hoverTimer);
        hoverTimer = setTimeout(() => {
            if (!isExpanded && !animationInProgress) {
                expandWindow();
            }
        }, 200); // 延遲 0.2 秒才展開
    });

    // 自動隱藏邏輯
    /*function hideWindow() {
        isExpanded = false;
        floatingWindow.style.width = '150px';
        toggleContainer.style.display = 'none';
        fileContainer.style.display = 'none';
        fileNameDisplay.style.display = 'none';
        dateControlContainer.style.display = 'none';
        storeControlContainer.style.display = 'none';
        actionButtonContainer.style.display = 'none';
        resultDisplay.style.display = 'none';
        byText.style.display = 'none';
    }*/

    function hideWindow() {
        if (animationInProgress) return;
        animationInProgress = true;
        isExpanded = false;
        resultDisplay.style.display = 'none';

        const containers = [
            toggleContainer,
            fileContainer,
            fileNameDisplay,
            dateControlContainer,
            storeControlContainer,
            actionButtonContainer,
            byText
        ];

        containers.reverse().forEach((container, index) => {
            setTimeout(() => {
                container.style.opacity = '0';
                container.style.transform = 'translateY(-10px)';
                setTimeout(() => {
                    container.style.display = 'none';
                    if (index === containers.length - 1) {
                        floatingWindow.style.width = '150px';
                        animationInProgress = false;
                    }
                }, 100); // 等待動畫結束再隱藏
            }, index * 50);
        });
    }

    // 滑鼠移出收合
    floatingWindow.addEventListener('mouseleave', () => {
        if (hoverTimer) clearTimeout(hoverTimer);
        hoverTimer = setTimeout(() => {
            if (isExpanded && isAutoHideEnabled && !animationInProgress) {
                hideWindow();
            }
        }, 200); // 延遲 0.2 秒才隱藏

    });

        // 自訂按鈕點擊事件
    fileButton.addEventListener('click', () => {
        fileInput.click();
    });

    // 上一天按鈕點擊事件
    prevDayButton.addEventListener('click', () => {
        if (!prevDayButton.disabled) {
            const currentDay = parseInt(dateSelect.value);
            if (currentDay > 1) {
                dateSelect.value = (currentDay - 1).toString();
            }
        }
    });

    // 下一天按鈕點擊事件
    nextDayButton.addEventListener('click', () => {
        if (!nextDayButton.disabled) {
            const currentDay = parseInt(dateSelect.value);
            const maxDay = dateSelect.options.length;
            if (currentDay < maxDay) {
                dateSelect.value = (currentDay + 1).toString();
            }
        }
    });

    // 輸入當日班表按鈕點擊事件
    inputDayButton.addEventListener('click', () => {
        if (!inputDayButton.disabled) {
            const selectedDate = dateSelect.value;
            if (!selectedDate || !Object.keys(Ximen).some(name => Ximen[name][selectedDate])) {
                alert(`日期 ${selectedDate} 無有效班表資料`);
                return;
            }
            const schedules = [];
            if (window.jQuery) {
                const $ = window.jQuery;
                let currentIndex = 0; // 從 X=0 開始
                for (const name of Object.keys(Ximen)) {
                    if (Ximen[name][selectedDate]?.value !== '休') {
                        const { value, S, E } = Ximen[name][selectedDate];
                        const valueStr = Array.isArray(value) ? value.join(',') : value;
                        schedules.push(`${name}: [${valueStr}, S="${S}", E="${E}"]`);

                        // 檢查是否有下一筆資料，決定是否新增元素
                        const hasNext = Object.keys(Ximen).some(n => n > name && Ximen[n][selectedDate]?.value !== '休');
                        if (hasNext) {
                            $('#add-table-shift').click(); // 新增新的 maids[X] 元素
                        }

                        // 偵測並設定對應的 maids[X] 元素
                        setTimeout(() => {
                            let foundMaid = false;
                            let newIndex = currentIndex;
                            while (!foundMaid) {
                                const $maidSelect = $(`select[name="maids[${newIndex}][maid]"]`);
                                if ($maidSelect.length > 0) {
                                    foundMaid = true;
                                    // 設定 maids[X][store] 為 storeSelect 選擇的文字
                                    const $storeSelect = $(`select[name="maids[${newIndex}][store]"]`);
                                    if ($storeSelect.length > 0) {
                                        const selectedStore = storeSelect.value;
                                        let storeMatch = null;
                                        // 檢查是否包含該選項
                                        $storeSelect.find('option').each((index, option) => {
                                            const optionText = $(option).text();
                                            if (optionText.includes(selectedStore)) {
                                                storeMatch = $(option).val();
                                                return false; // 終止 each 迴圈
                                            }
                                        });
                                        if (!storeMatch) {
                                            alert(`目前沒有該店家: ${selectedStore}`);
                                            return; // 中止當前迭代
                                        }
                                        // 設置選項
                                        $storeSelect.val(storeMatch).trigger('change');
                                    }
                                    // 設定 maids[X][maid] 根據 name 的 text
                                    let foundValue = null;
                                    $maidSelect.find('option').each((index, option) => {
                                        if ($(option).text().trim() === name) {
                                            foundValue = $(option).val();
                                            return false; // 終止 each 迴圈
                                        }
                                    });
                                    if (foundValue) {
                                        $maidSelect.val(foundValue).trigger('change');
                                    } else {
                                        console.warn(`在 maids[${newIndex}][maid] 中未找到與 "${name}" 對應的選項`);
                                    }
                                    // 設定 maids[X][start_time] 和 maids[X][end_time]
                                    $(`input[name="maids[${newIndex}][start_time]"]`).val(S);
                                    $(`input[name="maids[${newIndex}][end_time]"]`).val(E);
                                } else {
                                    newIndex++; // 遞增 X，繼續尋找
                                }
                            }
                            currentIndex = newIndex + 1;
                        }, 100);
                    }
                }
            } else {
                for (const name of Object.keys(Ximen)) {
                    if (Ximen[name][selectedDate]?.value !== '休') {
                        const { value, S, E } = Ximen[name][selectedDate];
                        const valueStr = Array.isArray(value) ? value.join(',') : value;
                        schedules.push(`${name}: [${valueStr}, S="${S}", E="${E}"]`);
                    }
                }
            }
            const message = schedules.length > 0
                ? `日期 ${selectedDate} 的班表：\n${schedules.join('\n')}`
                : `日期 ${selectedDate} 無班表`;
            //alert(message);
        }
    });

    // 清空當日班表按鈕點擊事件
    clearDayButton.addEventListener('click', () => {
        // 檢查 jQuery 是否可用
        if (window.jQuery) {
            const $ = window.jQuery;
            // 觸發 table-shift-remove 按鈕
            const removeButtons = document.querySelectorAll('.table-shift-remove');
            if (removeButtons.length === 0) {
                alert('無可清空的班表按鈕');
            } else {
                removeButtons.forEach(btn => btn.click());
            }
            // 清空 maids[0][store] 下拉選單
            $('select[name="maids[0][store]"]').val(null).trigger('change');
            $('select[name="maids[0][maid]"]').val(null).trigger('change');
            $('input[name="maids[0][start_time]"]').val("");
            $('input[name="maids[0][end_time]"]').val("");
        } else {
            alert('jQuery 未載入，無法清空下拉選單');
        }
    });

    // 啟用控制元件
    function enableControls() {
        prevDayButton.disabled = false;
        prevDayButton.style.cursor = 'pointer';
        prevDayButton.style.opacity = '1';
        prevDayButton.onmouseover = function() {
            this.style.backgroundColor = '#e0e0e0';
            this.style.boxShadow = '0 1px 3px rgba(0,0,0,0.1)';
        };
        prevDayButton.onmouseout = function() {
            this.style.backgroundColor = '#fff';
            this.style.boxShadow = 'none';
        };
        prevDayButton.onmousedown = function() {
            this.style.backgroundColor = '#d0d0d0';
            this.style.transform = 'scale(0.95)';
        };
        prevDayButton.onmouseup = function() {
            this.style.backgroundColor = '#e0e0e0';
            this.style.transform = 'scale(1)';
        };

        dateSelect.disabled = false;
        dateSelect.style.cursor = 'pointer';
        dateSelect.style.opacity = '1';

        nextDayButton.disabled = false;
        nextDayButton.style.cursor = 'pointer';
        nextDayButton.style.opacity = '1';
        nextDayButton.onmouseover = function() {
            this.style.backgroundColor = '#e0e0e0';
            this.style.boxShadow = '0 1px 3px rgba(0,0,0,0.1)';
        };
        nextDayButton.onmouseout = function() {
            this.style.backgroundColor = '#fff';
            this.style.boxShadow = 'none';
        };
        nextDayButton.onmousedown = function() {
            this.style.backgroundColor = '#d0d0d0';
            this.style.transform = 'scale(0.95)';
        };
        nextDayButton.onmouseup = function() {
            this.style.backgroundColor = '#e0e0e0';
            this.style.transform = 'scale(1)';
        };

        inputDayButton.disabled = false;
        inputDayButton.style.cursor = 'pointer';
        inputDayButton.style.opacity = '1';
        inputDayButton.onmouseover = function() {
            this.style.backgroundColor = '#e0e0e0';
            this.style.boxShadow = '0 1px 3px rgba(0,0,0,0.1)';
        };
        inputDayButton.onmouseout = function() {
            this.style.backgroundColor = '#fff';
            this.style.boxShadow = 'none';
        };
        inputDayButton.onmousedown = function() {
            this.style.backgroundColor = '#d0d0d0';
            this.style.transform = 'scale(0.95)';
        };
        inputDayButton.onmouseup = function() {
            this.style.backgroundColor = '#e0e0e0';
            this.style.transform = 'scale(1)';
        };
    }

    // 禁用控制元件
    function disableControls() {
        prevDayButton.disabled = true;
        prevDayButton.style.cursor = 'not-allowed';
        prevDayButton.style.opacity = '0.5';
        prevDayButton.onmouseover = null;
        prevDayButton.onmouseout = null;
        prevDayButton.onmousedown = null;
        prevDayButton.onmouseup = null;

        dateSelect.disabled = true;
        dateSelect.style.cursor = 'not-allowed';
        dateSelect.style.opacity = '0.5';
        dateSelect.innerHTML = '';

        nextDayButton.disabled = true;
        nextDayButton.style.cursor = 'not-allowed';
        nextDayButton.style.opacity = '0.5';
        nextDayButton.onmouseover = null;
        nextDayButton.onmouseout = null;
        nextDayButton.onmousedown = null;
        nextDayButton.onmouseup = null;

        inputDayButton.disabled = true;
        inputDayButton.style.cursor = 'not-allowed';
        inputDayButton.style.opacity = '0.5';
        inputDayButton.onmouseover = null;
        inputDayButton.onmouseout = null;
        inputDayButton.onmousedown = null;
        inputDayButton.onmouseup = null;
    }

    // 檔案選擇事件
    fileInput.addEventListener('change', (event) => {
        const file = event.target.files[0];
        if (file) {
            fileNameDisplay.textContent = file.name;

            // 等待 SheetJS 載入
            let attempts = 0;
            const maxAttempts = 100; // 最多等待 10 秒 (100 * 100ms)
            const checkSheetJS = () => {
                if (window.XLSX && sheetJSLoaded) {
                    const reader = new FileReader();
                    reader.onload = function(e) {
                        try {
                            const data = new Uint8Array(e.target.result);
                            const workbook = XLSX.read(data, { type: 'array' });
                            const sheet = workbook.Sheets[workbook.SheetNames[0]];
                            const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });

                            Ximen = {};
                            ShiftTimes = {};
                            let dateColumns = [];
                            let shiftColumnIndex = -1;
                            let isInDataSection = false;
                            let isInShiftSection = true;
                            let rowCount = 0;

                            let validShifts = [];

                            // 第一階段：解析班別和員工資料
                            rows.forEach((row, index) => {
                                // 檢查第一行，找出日期欄位和班別欄位
                                if (index === 0) {
                                    if (row[2] === '日期') {
                                        dateColumns = row.slice(3).filter(cell => /^\d+$/.test(cell));
                                        shiftColumnIndex = row.indexOf('班別');
                                    }
                                    return;
                                }

                                // 處理班別時間（在第一個工作表，班別欄位之後）
                                if (isInShiftSection && shiftColumnIndex !== -1 && row[shiftColumnIndex]) {
                                    const shift = String(row[shiftColumnIndex]).trim();
                                    if (shift === '總時數') {
                                        isInShiftSection = false;
                                    } else {
                                        validShifts.push(shift);
                                        let timeRange = row[shiftColumnIndex + 2]; // 時間欄位在班別後第2欄
                                        // 處理換行符號，保留換行前內容
                                        const newlineIndex = timeRange.indexOf('\n');
                                        if (newlineIndex !== -1) {
                                            timeRange = timeRange.substring(0, newlineIndex).trim();
                                        }
                                        if (timeRange && timeRange.includes('~')) {
                                            const [start, end] = timeRange.split('~').map(t => t.trim());
                                            ShiftTimes[shift] = [start, end];
                                        }
                                    }
                                }

                                // 檢查是否進入資料區（從 "名稱" 開始）
                                if (row[0] === '名稱') {
                                    isInDataSection = true;
                                    return;
                                }

                                // 檢查是否結束資料區（遇到 "總和"）
                                if (row[0] === '總和') {
                                    isInDataSection = false;
                                    isInShiftSection = false;
                                    return;
                                }

                                // 處理員工班表（僅設置 value）
                                if (isInDataSection && String(row[0] || '').trim()) {
                                    const name = String(row[0]).trim();
                                    rowCount++;
                                    Ximen[name] = {};
                                    dateColumns.forEach((date, i) => {
                                        let value = String(row[3 + i] || '休').trim();
                                        let shifts;
                                        // 檢查是否包含「創」、「北」、「信」、「公」、「西」
                                        if (['創', '北', '信', '公', '西'].some(prefix => value.includes(prefix))) {
                                            shifts = '休';
                                        } else {
                                            // 處理班別為陣列或字串
                                            let tempShifts = [];
                                            let remaining = value.replace(/\+/g, '');
                                            while (remaining) {
                                                let matched = false;
                                                for (const shift of validShifts) {
                                                    if (remaining.startsWith(shift)) {
                                                        tempShifts.push(shift);
                                                        remaining = remaining.slice(shift.length).trim();
                                                        matched = true;
                                                        break;
                                                    }
                                                }
                                                if (!matched) {
                                                    remaining = '';
                                                }
                                            }
                                            shifts = tempShifts.length > 0 ? (tempShifts.length === 1 ? tempShifts[0] : tempShifts) : '休';
                                        }
                                        Ximen[name][date] = { value: shifts, S: '', E: '' };
                                    });
                                }
                            });

                            // 第二階段：設置 S 和 E
                            Object.keys(Ximen).forEach(name => {
                                Object.keys(Ximen[name]).forEach(date => {
                                    const { value } = Ximen[name][date];
                                    let S = '', E = '';
                                    if (value !== '休') {
                                        if (typeof value === 'string') {
                                            if (ShiftTimes[value]) {
                                                S = ShiftTimes[value][0];
                                                E = ShiftTimes[value][1];
                                            }
                                        } else if (Array.isArray(value)) {
                                            if (ShiftTimes[value[0]]) {
                                                S = ShiftTimes[value[0]][0];
                                            }
                                            if (value[1] && ShiftTimes[value[1]]) {
                                                E = ShiftTimes[value[1]][1];
                                            } else if (ShiftTimes[value[0]]) {
                                                E = ShiftTimes[value[0]][1];
                                            }
                                        }
                                    }
                                    Ximen[name][date].S = S;
                                    Ximen[name][date].E = E;
                                });
                            });

                            // 動態生成日期下拉選單
                            dateSelect.innerHTML = '';
                            dateColumns.forEach(day => {
                                const option = document.createElement('option');
                                option.value = day;
                                option.textContent = day;
                                dateSelect.appendChild(option);
                            });
                            dateSelect.value = '1'; // 預設選擇第 1 天
                            enableControls();

                            // 顯示解析結果（所有員工，前 3 個日期，顯示 value 和時間）
                            const employeePreview = Object.keys(Ximen).map(name => {
                                const dates = Object.keys(Ximen[name]).slice(0, 15);
                                return `${name}: ${dates.map(date => {
                                    const { value, S, E } = Ximen[name][date];
                                    const valueStr = Array.isArray(value) ? value.join(',') : value;
                                    return `${date}=[${valueStr}, S="${S}", E="${E}"]`;
                                }).join(', ')}`;
                            }).join('\n');
                            const shiftPreview = Object.keys(ShiftTimes).map(shift => {
                                return `${shift}: ${ShiftTimes[shift][0]}~${ShiftTimes[shift][1]}`;
                            }).join('\n');
                            resultDisplay.textContent = `員工數: ${rowCount}\n員工班表預覽:\n${employeePreview || '無數據'}\n\n班別時間:\n${shiftPreview || '無班別資料'}`;
                        } catch (error) {
                            resultDisplay.textContent = '解析 Excel 檔案失敗: ' + error.message;
                        }
                    };
                    reader.onerror = function() {
                        resultDisplay.textContent = '讀取 Excel 檔案失敗';
                    };
                    reader.readAsArrayBuffer(file);
                } else if (attempts < maxAttempts) {
                    attempts++;
                    setTimeout(checkSheetJS, 100);
                } else {
                    resultDisplay.textContent = 'SheetJS 庫載入超時，請檢查網路、禁用廣告攔截器或重新整理頁面';
                }
            };
            checkSheetJS();
        } else {
            fileNameDisplay.textContent = '尚未選擇檔案';
            resultDisplay.textContent = '無解析結果';
        }
    });
})();