document.addEventListener('DOMContentLoaded', () => {
    const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/AKfycbxfEsIAxWNXP8zynB0aPHXeWVmMVvCPvwolvdYQ9QWMbXnZxMqhOX75R_R0XSuSX-uVpw/exec';
    const ATTENDEES_STORAGE_KEY = 'aiWorkshopAttendees_v2';
    let allAttendees = [];
    const GROUPS = ["第一組", "第二組", "第三組", "第四組", "第五組", "第六組", "第七組"];
    const tableBody = document.getElementById('attendeeTableBody');
    const searchInput = document.getElementById('searchInput');
    const addAttendeeBtn = document.getElementById('addAttendeeBtn');
    const exportBtn = document.getElementById('exportBtn');
    const loadingOverlay = document.getElementById('loading-overlay');
    const loadingMessage = document.getElementById('loading-message');
    const modal = document.getElementById('attendeeModal');
    const modalContent = document.getElementById('modal-content');
    const modalTitle = document.getElementById('modalTitle');
    const closeModalBtn = document.getElementById('modal-close');
    const cancelModalBtn = document.getElementById('formCancel');
    const overlay = document.getElementById('modal-overlay');
    const saveBtn = document.getElementById('formSave');
    const attendeeForm = document.getElementById('attendeeForm');
    const formGroup = document.getElementById('formGroup');
    const formName = document.getElementById('formName');
    const formOrganization = document.getElementById('formOrganization');
    const formTitle = document.getElementById('formTitle');
    const formCategory = document.getElementById('formCategory');
    const editUuidHidden = document.getElementById('editUuid');
    const alertModal = document.getElementById('alertModal');
    const alertTitle = document.getElementById('alertTitle');
    const alertIdBadge = document.getElementById('alertIdBadge');
    const alertMessage = document.getElementById('alertMessage');
    const alertCloseBtn = document.getElementById('alertCloseBtn');
    const confirmModal = document.getElementById('confirmModal');
    const confirmTitle = document.getElementById('confirmTitle');
    const confirmMessage = document.getElementById('confirmMessage');
    const confirmCancelBtn = document.getElementById('confirmCancelBtn');
    const confirmOkBtn = document.getElementById('confirmOkBtn');
    const errorModal = document.getElementById('errorModal');
    const errorTitle = document.getElementById('errorTitle');
    const errorMessage = document.getElementById('errorMessage');
    const errorCloseBtn = document.getElementById('errorCloseBtn');
    let confirmCallback = null;
    const toastEl = document.getElementById('toast');
    const toastMessageEl = document.getElementById('toast-message');
    const clockEl = document.getElementById('realtime-clock');

    function saveData() {
        try {
            localStorage.setItem(ATTENDEES_STORAGE_KEY, JSON.stringify(allAttendees));
        } catch (e) {
            console.error("儲存到 localStorage 失敗:", e);
            openErrorModal("儲存失敗", "無法將簽到進度儲存到瀏覽器。您的簽到資料在重新整理後可能會遺失。");
        }
    }

    async function loadData() {
        const storedAttendees = localStorage.getItem(ATTENDEES_STORAGE_KEY);
        if (storedAttendees) {
            try {
                allAttendees = JSON.parse(storedAttendees);
                loadingMessage.textContent = "已載入儲存的簽到進度...";
                await new Promise(resolve => setTimeout(resolve, 500));
            } catch (e) {
                console.error("解析 localStorage 資料失敗:", e);
                localStorage.removeItem(ATTENDEES_STORAGE_KEY);
            }
        }
        if (allAttendees.length === 0) {
            try {
                loadingMessage.textContent = "正在從 Google Sheet 載入學員名單...";
                const response = await fetch(APPS_SCRIPT_URL);
                if (!response.ok) throw new Error(`網路回應錯誤: ${response.status} ${response.statusText}`);
                const data = await response.json();
                if (data && Array.isArray(data)) {
                    allAttendees = data.map(item => ({
                        uuid: item.uuid || crypto.randomUUID(),
                        group: item.group || '',
                        name: item.name || '',
                        category: item.category || '學員',
                        organization: item.organization || '',
                        title: item.title || '',
                        checkedIn: item.checkedIn || false,
                        checkInTime: item.checkInTime || null
                    }));
                    saveData();
                } else if (data && data.error) {
                    throw new Error(`Apps Script 錯誤: ${data.message}`);
                } else {
                    throw new Error("從 Google Sheet 載入的資料格式不正確，並非陣列。");
                }
            } catch (e) {
                console.error("載入 Google Sheet 資料失敗:", e);
                loadingMessage.textContent = "載入學員名單失敗。";
                openErrorModal("載入失敗", `無法從 Google Sheet 取得資料。<br><br>錯誤訊息: ${e.message}<br><br>請檢查網路連線及 Apps Script 部署狀態。`);
                loadingOverlay.classList.add('opacity-0', 'pointer-events-none');
                return;
            }
        }
        filterAndRender();
        updateStats();
        loadingOverlay.classList.add('opacity-0', 'pointer-events-none');
    }

    function updateClock() {
        const now = new Date();
        const timeString = now.toLocaleTimeString('zh-TW', { hour12: false });
        const dateString = now.toLocaleDateString('zh-TW');
        clockEl.textContent = `${dateString} ${timeString}`;
    }

    function showToast(message) {
        toastMessageEl.textContent = message;
        toastEl.classList.add('toast-show');
        setTimeout(() => {
            toastEl.classList.remove('toast-show');
        }, 3000);
    }

    function openModal() { modal.classList.remove('opacity-0', 'pointer-events-none'); modalContent.classList.remove('-translate-y-10'); }
    function closeModal() { modal.classList.add('opacity-0', 'pointer-events-none'); modalContent.classList.add('-translate-y-10'); attendeeForm.reset(); editUuidHidden.value = ''; }

    function openAlertModal(person) {
        alertTitle.innerHTML = `<span class="font-normal">${person.group}</span> ${person.name}`;
        alertIdBadge.innerHTML = `<p class="flex items-center justify-center gap-2"><ion-icon name="person-outline"></ion-icon> ${person.title || '學員'}</p>`;
        if (person.group === '第六組' || person.group === '第七組') {
            alertMessage.innerHTML = '<p>一切順利，歡迎您的蒞臨！</p><p class="mt-3 font-semibold" style="color: #E89152;"><ion-icon name="location-outline"></ion-icon> 上課地點在貓頭鷹手作教室唷</p>';
        } else {
            alertMessage.innerHTML = '<p>一切順利，歡迎您的蒞臨！</p>';
        }
        alertModal.classList.remove('opacity-0', 'pointer-events-none');
        alertModal.querySelector('.modal-content').classList.remove('-translate-y-10');
    }
    function closeAlertModal() { alertModal.classList.add('opacity-0', 'pointer-events-none'); alertModal.querySelector('.modal-content').classList.add('-translate-y-10'); }

    function openConfirmModal(title, message, callback) {
        confirmTitle.textContent = title;
        confirmMessage.textContent = message;
        confirmCallback = callback;
        confirmModal.classList.remove('opacity-0', 'pointer-events-none');
        confirmModal.querySelector('.modal-content').classList.remove('-translate-y-10');
    }
    function closeConfirmModal() { confirmModal.classList.add('opacity-0', 'pointer-events-none'); confirmModal.querySelector('.modal-content').classList.add('-translate-y-10'); confirmCallback = null; }

    function openErrorModal(title, message) {
        errorTitle.textContent = title;
        errorMessage.innerHTML = message;
        errorModal.classList.remove('opacity-0', 'pointer-events-none');
        errorModal.querySelector('.modal-content').classList.remove('-translate-y-10');
    }
    function closeErrorModal() { errorModal.classList.add('opacity-0', 'pointer-events-none'); errorModal.querySelector('.modal-content').classList.add('-translate-y-10'); }

    function animateValue(element, start, end, duration, isPercentage = false) {
        if (!element) return;
        if (start === end) {
            element.textContent = isPercentage ? `${end}%` : end;
            return;
        }
        let startTimestamp = null;
        const step = (timestamp) => {
            if (!startTimestamp) startTimestamp = timestamp;
            const progress = Math.min((timestamp - startTimestamp) / duration, 1);
            const currentValue = Math.floor(progress * (end - start) + start);
            element.textContent = isPercentage ? `${currentValue}%` : currentValue;
            if (progress < 1) {
                window.requestAnimationFrame(step);
            } else {
                element.textContent = isPercentage ? `${end}%` : end;
                element.classList.add('number-pop-animation');
                setTimeout(() => { element.classList.remove('number-pop-animation'); }, 300);
            }
        };
        window.requestAnimationFrame(step);
    }

    function renderTable(data) {
        tableBody.innerHTML = '';
        if (data.length === 0) {
            tableBody.innerHTML = '<tr><td colspan="7" class="text-center p-8 text-gray-500">找不到資料</td></tr>';
            return;
        }
        const groupSortOrder = GROUPS.reduce((acc, group, index) => {
            acc[group] = index;
            return acc;
        }, {});
        const sortCompare = (a, b) => {
            const groupA = groupSortOrder[a.group] ?? 99;
            const groupB = groupSortOrder[b.group] ?? 99;
            if (groupA !== groupB) return groupA - groupB;
            return a.name.localeCompare(b.name, 'zh-Hant');
        };
        const checkedInList = data.filter(p => p.checkedIn);
        const notCheckedInList = data.filter(p => !p.checkedIn);
        checkedInList.sort((a, b) => {
            const groupA = groupSortOrder[a.group] ?? 99;
            const groupB = groupSortOrder[b.group] ?? 99;
            if (groupA !== groupB) return groupA - groupB;
            return new Date(b.checkInTime) - new Date(a.checkInTime);
        });
        notCheckedInList.sort(sortCompare);
        const combinedList = [...checkedInList, ...notCheckedInList];
        combinedList.forEach(p => {
            const row = document.createElement('tr');
            row.className = p.checkedIn ? 'bg-green-50' : '';
            row.innerHTML = `
                <td class="table-cell"><span class="flex items-center gap-2 ${p.checkedIn ? 'text-green-600' : 'text-gray-500'}"><ion-icon name="${p.checkedIn ? 'checkmark-circle' : 'ellipse-outline'}"></ion-icon>${p.checkedIn ? '已報到' : '未報到'}</span></td>
                <td class="table-cell font-mono">${p.group}</td>
                <td class="table-cell"><span class="font-semibold ${p.category === '講師' ? 'text-purple-700' : (p.category === '工作人員' ? 'text-orange-700' : 'text-gray-800')}">${p.category}</span></td>
                <td class="table-cell">${p.organization}</td>
                <td class="table-cell">${p.title}</td>
                <td class="table-cell font-semibold">${p.name}</td>
                <td class="table-cell">
                    <div class="flex items-center justify-end gap-1">
                        ${p.checkedIn ? `
                            <div class="flex-grow text-left text-green-600">
                                <p class="font-semibold flex items-center gap-1"><ion-icon name="checkmark-done-outline"></ion-icon>已簽到</p>
                                <p class="text-xs text-gray-500">${p.checkInTime ? new Date(p.checkInTime).toLocaleString('zh-TW', { hour: '2-digit', minute: '2-digit', second: '2-digit' }) : ''}</p>
                            </div>
                            <button class="cancel-btn text-gray-500 p-2 rounded-md hover:bg-red-100 hover:text-red-600" data-uuid="${p.uuid}" title="取消簽到">
                                <ion-icon name="arrow-undo-outline" class="pointer-events-none"></ion-icon>
                            </button>`
                        : `
                            <button class="checkin-btn text-white px-3 py-1 rounded-md" style="background-color: #5B9FB6;" data-uuid="${p.uuid}" onmouseover="this.style.backgroundColor='#2B4856'" onmouseout="this.style.backgroundColor='#5B9FB6'">簽到</button>`
                        }
                        <button class="edit-btn text-gray-500 p-2 rounded-md hover:bg-gray-200" data-uuid="${p.uuid}" title="編輯資料">
                            <ion-icon name="create-outline" class="pointer-events-none"></ion-icon>
                        </button>
                    </div>
                </td>`;
            tableBody.appendChild(row);
        });
    }

    function updateStats() {
        const total = allAttendees.length;
        const newCheckedIn = allAttendees.filter(a => a.checkedIn).length;
        const newRate = total > 0 ? Math.round((newCheckedIn / total) * 100) : 0;
        const totalAttendeesEl = document.getElementById('totalAttendees');
        const totalCheckedInEl = document.getElementById('totalCheckedIn');
        const totalAttendanceRateEl = document.getElementById('totalAttendanceRate');
        animateValue(totalAttendeesEl, parseInt(totalAttendeesEl.textContent) || 0, total, 300);
        animateValue(totalCheckedInEl, parseInt(totalCheckedInEl.textContent) || 0, newCheckedIn, 300);
        animateValue(totalAttendanceRateEl, parseInt(totalAttendanceRateEl.textContent.replace('%','')) || 0, newRate, 300, true);
        const groupStatsContainer = document.getElementById('group-stats-container');
        groupStatsContainer.innerHTML = '';
        const studentAttendees = allAttendees.filter(a => a.category === '學員');
        GROUPS.forEach(groupName => {
            const groupAttendees = studentAttendees.filter(a => a.group === groupName);
            const totalInGroup = groupAttendees.length;
            const checkedInInGroup = groupAttendees.filter(a => a.checkedIn).length;
            const card = document.createElement('div');
            const isComplete = checkedInInGroup === totalInGroup && totalInGroup > 0;
            card.className = 'p-3 rounded-xl text-center';
            card.style.backgroundColor = isComplete ? '#5B9FB6' : '#A8C5D1';
            card.innerHTML = `
                <h3 class="text-sm font-semibold" style="color: ${isComplete ? '#FFFFFF' : '#2B4856'};">${groupName}</h3>
                <p class="text-3xl font-bold mt-1" style="color: ${isComplete ? '#FFFFFF' : '#1A2328'};">
                    <span id="group-checked-${groupName}">${checkedInInGroup}</span>
                    <span class="text-xl" style="color: ${isComplete ? 'rgba(255,255,255,0.7)' : '#5B9FB6'};">/ ${totalInGroup}</span>
                </p>
            `;
            groupStatsContainer.appendChild(card);
        });
    }

    function handleCheckIn(uuid) {
        const person = allAttendees.find(p => p.uuid === uuid);
        if (person && !person.checkedIn) {
            person.checkedIn = true;
            person.checkInTime = new Date().toISOString();
            openAlertModal(person);
            filterAndRender();
            updateStats();
            saveData();
        }
    }

    function handleCancelCheckIn(uuid) {
        const person = allAttendees.find(p => p.uuid === uuid);
        if (!person) return;
        openConfirmModal('取消簽到確認', `您確定要取消【${person.name}】的簽到狀態嗎？`, () => {
            if (person.checkedIn) {
                person.checkedIn = false;
                person.checkInTime = null;
                filterAndRender();
                updateStats();
                saveData();
            }
        });
    }

    function handleEdit(uuid) {
        const person = allAttendees.find(p => p.uuid === uuid);
        if (person) {
            modalTitle.textContent = '編輯參與者資料';
            editUuidHidden.value = person.uuid;
            formGroup.value = person.group;
            formName.value = person.name;
            formOrganization.value = person.organization;
            formTitle.value = person.title;
            formCategory.value = person.category;
            openModal();
        }
    }

    function filterAndRender() {
        const searchTerm = searchInput.value.toLowerCase();
        const filtered = allAttendees.filter(p =>
            p.name.toLowerCase().includes(searchTerm) ||
            p.organization.toLowerCase().includes(searchTerm) ||
            p.title.toLowerCase().includes(searchTerm) ||
            p.group.toLowerCase().includes(searchTerm)
        );
        renderTable(filtered);
    }

    function exportToCSV() {
        const headers = ['狀態', '組別', '身份類別', '姓名', '服務單位/部門', '職稱', '報到時間'];
        const rows = allAttendees.map(p => [
            p.checkedIn ? '已報到' : '未報到',
            p.group,
            p.category,
            p.name,
            p.organization,
            p.title,
            p.checkInTime ? new Date(p.checkInTime).toLocaleString('zh-TW') : ''
        ].map(field => `"${String(field).replace(/"/g, '""')}"`).join(','));
        const csvContent = "data:text/csv;charset=utf-8,\uFEFF" + [headers.join(','), ...rows].join('\n');
        const encodedUri = encodeURI(csvContent);
        const link = document.createElement("a");
        link.setAttribute("href", encodedUri);
        link.setAttribute("download", `AI工作坊報到資料_${new Date().toISOString().slice(0,10)}.csv`);
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
    }

    addAttendeeBtn.addEventListener('click', () => {
        modalTitle.textContent = '新增參與者';
        attendeeForm.reset();
        editUuidHidden.value = '';
        openModal();
    });

    closeModalBtn.addEventListener('click', closeModal);
    cancelModalBtn.addEventListener('click', closeModal);
    overlay.addEventListener('click', closeModal);
    alertCloseBtn.addEventListener('click', closeAlertModal);
    confirmCancelBtn.addEventListener('click', closeConfirmModal);
    confirmOkBtn.addEventListener('click', () => {
        if (typeof confirmCallback === 'function') {
            confirmCallback();
        }
        closeConfirmModal();
    });
    errorCloseBtn.addEventListener('click', closeErrorModal);
    exportBtn.addEventListener('click', exportToCSV);

    saveBtn.addEventListener('click', () => {
        const uuidToSave = editUuidHidden.value;
        const data = {
            group: formGroup.value.trim(),
            name: formName.value.trim(),
            organization: formOrganization.value.trim(),
            title: formTitle.value.trim(),
            category: formCategory.value,
        };
        if (!data.name) {
            openErrorModal("資料不完整", "「姓名」為必填欄位！");
            return;
        }
        if (uuidToSave) {
            const index = allAttendees.findIndex(p => p.uuid === uuidToSave);
            if (index !== -1) {
                const existingData = allAttendees[index];
                allAttendees[index] = { ...existingData, ...data };
            }
        } else {
            allAttendees.unshift({
                ...data,
                uuid: crypto.randomUUID(),
                checkedIn: false,
                checkInTime: null
            });
        }
        closeModal();
        filterAndRender();
        updateStats();
        saveData();
        showToast('資料已儲存！');
    });

    searchInput.addEventListener('input', filterAndRender);

    tableBody.addEventListener('click', (e) => {
        const checkinBtn = e.target.closest('.checkin-btn');
        const editBtn = e.target.closest('.edit-btn');
        const cancelBtn = e.target.closest('.cancel-btn');
        if (checkinBtn) handleCheckIn(checkinBtn.dataset.uuid);
        if (editBtn) handleEdit(editBtn.dataset.uuid);
        if (cancelBtn) handleCancelCheckIn(cancelBtn.dataset.uuid);
    });

    loadData();
    updateClock();
    setInterval(updateClock, 1000);
});
