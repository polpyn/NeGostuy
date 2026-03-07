// ============================================================================
// КОНФИГУРАЦИЯ
// ============================================================================

const CONFIG = {
    API_URL: '/api/upload/',
    HEALTH_URL: '/api/health/',
    DOWNLOAD_URL: '/api/download/',
    MAX_FILE_SIZE: 16 * 1024 * 1024,
    ALLOWED_EXTENSION: '.docx'
};

// ============================================================================
// DOM ЭЛЕМЕНТЫ
// ============================================================================

const elements = {
    fileInput: null,
    templateInput: null,
    submitBtn: null,
    fileName: null,
    templateName: null,
    errorMessage: null,
    resultText: null,
    loadingIndicator: null,
    analysisResults: null
};

// Выбранные файлы
let selectedFile = null;
let selectedTemplate = null;

// ============================================================================
// ИНИЦИАЛИЗАЦИЯ
// ============================================================================

document.addEventListener('DOMContentLoaded', () => {
    console.log('🚀 НеГОСТуй - инициализация');

    // Находим элементы
    elements.fileInput = document.getElementById('fileInput');
    elements.templateInput = document.getElementById('templateInput');
    elements.submitBtn = document.getElementById('submitBtn');
    elements.fileName = document.getElementById('fileName');
    elements.templateName = document.getElementById('templateName');
    elements.errorMessage = document.getElementById('errorMessage');
    elements.resultText = document.getElementById('resultText');
    elements.loadingIndicator = document.getElementById('loadingIndicator');
    elements.analysisResults = document.getElementById('analysisResults');

    // Проверка сервера
    checkServerHealth();

    // Обработчики
    if (elements.fileInput) {
        elements.fileInput.addEventListener('change', handleFileSelect);
    }
    if (elements.templateInput) {
        elements.templateInput.addEventListener('change', handleTemplateSelect);
    }
    if (elements.submitBtn) {
        elements.submitBtn.addEventListener('click', handleSubmit);
    }

    console.log('✅ Приложение готово');
});

// ============================================================================
// ПРОВЕРКА СЕРВЕРА
// ============================================================================

async function checkServerHealth() {
    try {
        const response = await fetch(CONFIG.HEALTH_URL);
        const data = await response.json();
        if (data.status === 'ok') {
            console.log('✅ Сервер доступен');
        }
    } catch (error) {
        console.warn('⚠️ Сервер недоступен');
    }
}

// ============================================================================
// ВЫБОР ДОКУМЕНТА
// ============================================================================

function handleFileSelect(event) {
    const file = event.target.files[0];
    clearMessages();

    if (!file) return;

    console.log('📁 Документ:', file.name);

    const error = validateFile(file);
    if (error) {
        showError(error);
        elements.fileInput.value = '';
        selectedFile = null;
        updateSubmitButton();
        return;
    }

    selectedFile = file;
    elements.fileName.textContent = `✓ ${file.name}`;
    elements.fileName.style.display = 'block';

    updateSubmitButton();
}

// ============================================================================
// ВЫБОР РАМКИ
// ============================================================================

function handleTemplateSelect(event) {
    const file = event.target.files[0];

    if (!file) return;

    console.log('🖼️ Рамка:', file.name);

    if (!file.name.endsWith('.docx')) {
        showError('Рамка должна быть в формате .docx');
        elements.templateInput.value = '';
        selectedTemplate = null;
        return;
    }

    selectedTemplate = file;
    elements.templateName.textContent = `✓ Рамка: ${file.name}`;
    elements.templateName.style.display = 'block';
}

// ============================================================================
// КНОПКА ОТПРАВКИ
// ============================================================================

function updateSubmitButton() {
    if (selectedFile && elements.submitBtn) {
        elements.submitBtn.style.display = 'block';
    } else if (elements.submitBtn) {
        elements.submitBtn.style.display = 'none';
    }
}

function handleSubmit() {
    if (!selectedFile) {
        showError('Сначала выберите документ');
        return;
    }
    uploadFile(selectedFile, selectedTemplate);
}

// ============================================================================
// ВАЛИДАЦИЯ
// ============================================================================

function validateFile(file) {
    const extension = '.' + file.name.split('.').pop().toLowerCase();
    if (extension !== CONFIG.ALLOWED_EXTENSION) {
        return 'Разрешены только файлы .docx';
    }
    if (file.size > CONFIG.MAX_FILE_SIZE) {
        return 'Файл слишком большой. Максимум 16 МБ';
    }
    if (file.size === 0) {
        return 'Файл пустой';
    }
    return null;
}

// ============================================================================
// ЗАГРУЗКА НА СЕРВЕР
// ============================================================================

async function uploadFile(file, template) {
    showLoading();

    const formData = new FormData();
    formData.append('file', file);

    // Добавляем рамку если выбрана
    if (template) {
        formData.append('template', template);
        console.log('📤 Отправка с рамкой:', template.name);
    } else {
        console.log('📤 Отправка без рамки');
    }

    try {
        const response = await fetch(CONFIG.API_URL, {
            method: 'POST',
            body: formData
        });

        console.log('📥 Ответ:', response.status);

        if (!response.ok) {
            let errorText = 'Ошибка сервера';
            try {
                const errorData = await response.json();
                errorText = errorData.error || errorText;
            } catch (e) {
                errorText = `HTTP ${response.status}`;
            }
            throw new Error(errorText);
        }

        const data = await response.json();
        console.log('✅ Данные:', data);

        if (!data.success) {
            throw new Error(data.error || 'Ошибка обработки');
        }

        displayResults(data);

    } catch (error) {
        hideLoading();
        console.error('❌', error);

        if (error.message.includes('fetch')) {
            showError('Сервер недоступен. Запустите: python manage.py runserver');
        } else {
            showError(error.message);
        }
    }
}

// ============================================================================
// ОТОБРАЖЕНИЕ РЕЗУЛЬТАТОВ
// ============================================================================

function displayResults(data) {
    hideLoading();
    if (elements.resultText) elements.resultText.style.display = 'none';

    const { summary, elements: items } = data;

    let html = `
        <div class="result-summary">
            <h3>${data.document_name}</h3>

            <div class="summary-grid">
                <div class="summary-item">
                    <span>Всего элементов:</span>
                    <span>${summary.total_elements}</span>
                </div>
                <div class="summary-item">
                    <span>Заголовков:</span>
                    <span>${summary.headings}</span>
                </div>
                <div class="summary-item">
                    <span>Списков:</span>
                    <span>${summary.lists}</span>
                </div>
                <div class="summary-item">
                    <span>Текста:</span>
                    <span>${summary.texts}</span>
                </div>
                <div class="summary-item">
                    <span>Ошибок:</span>
                    <span>${summary.errors_count}</span>
                </div>
                <div class="summary-item">
                    <span>Предупреждений:</span>
                    <span>${summary.warnings_count}</span>
                </div>
            </div>

            ${summary.has_template ? `
                <p style="color: #4CAF50; margin-top: 8px; font-size: 13px;">
                    ✅ Рамка применена
                </p>
            ` : ''}

            <div class="grade-badge ${getGradeClass(summary.grade)}">
                ${summary.grade}
            </div>

            ${data.document_id ? `
                <button class="download-btn" onclick="downloadFile(${data.document_id})">
                    📥 Скачать исправленный документ
                </button>
            ` : ''}
        </div>

        <div class="elements-list">
            <h4>Детальный анализ (${items.length} элементов):</h4>
    `;

    items.forEach((item, index) => {
        html += createElementCard(item, index);
    });

    html += '</div>';

    elements.analysisResults.innerHTML = html;
    elements.analysisResults.style.display = 'block';
}

// ============================================================================
// КАРТОЧКА ЭЛЕМЕНТА
// ============================================================================

function createElementCard(item, index) {
    const typeLabels = {
        'heading': '📌 Заголовок',
        'text': '📝 Текст',
        'list_item': '📋 Список'
    };

    const statusIcons = {
        'correct': '✅',
        'warning': '⚠️',
        'error': '❌'
    };

    return `
        <div class="element-card status-${item.status}">
            <div class="element-header">
                <span style="font-weight: bold;">#${index + 1}</span>
                <span class="element-type">${typeLabels[item.type] || item.type}</span>
                <span>${statusIcons[item.status] || ''}</span>
            </div>
            <div class="element-text">"${item.text}"</div>
            ${item.errors && item.errors.length > 0 ? `
                <div class="element-errors">
                    <strong>Ошибки:</strong>
                    ${item.errors.map(err => `<p>${err}</p>`).join('')}
                </div>
            ` : ''}
            ${item.warnings && item.warnings.length > 0 ? `
                <div class="element-warnings">
                    <strong>Предупреждения:</strong>
                    ${item.warnings.map(w => `<p>${w}</p>`).join('')}
                </div>
            ` : ''}
            ${item.status === 'correct' ? `
                <p style="color: #4CAF50; margin-top: 8px; font-weight: bold;">✓ Соответствует ГОСТ</p>
            ` : ''}
        </div>
    `;
}

// ============================================================================
// СКАЧИВАНИЕ
// ============================================================================

function downloadFile(documentId) {
    const url = CONFIG.DOWNLOAD_URL + documentId + '/';
    console.log('📥 Скачивание:', url);

    fetch(url)
        .then(response => {
            if (!response.ok) throw new Error('Ошибка скачивания');
            return response.blob();
        })
        .then(blob => {
            const link = document.createElement('a');
            link.href = URL.createObjectURL(blob);
            link.download = 'GOST_документ.docx';
            document.body.appendChild(link);
            link.click();
            document.body.removeChild(link);
            URL.revokeObjectURL(link.href);
            console.log('✅ Скачан');
        })
        .catch(error => {
            console.error('❌', error);
            alert('Ошибка скачивания: ' + error.message);
        });
}

// ============================================================================
// ВСПОМОГАТЕЛЬНЫЕ
// ============================================================================

function getGradeClass(grade) {
    if (grade.includes('ИДЕАЛЬНО') || grade.includes('ОТЛИЧНО')) return 'grade-perfect';
    if (grade.includes('ХОРОШО') || grade.includes('ПРЕДУПРЕЖДЕНИЯ')) return 'grade-warning';
    return 'grade-error';
}

function showLoading() {
    if (elements.resultText) elements.resultText.style.display = 'none';
    if (elements.analysisResults) elements.analysisResults.style.display = 'none';
    if (elements.loadingIndicator) elements.loadingIndicator.style.display = 'block';
}

function hideLoading() {
    if (elements.loadingIndicator) elements.loadingIndicator.style.display = 'none';
}

function showError(message) {
    if (elements.errorMessage) {
        elements.errorMessage.textContent = '⚠️ ' + message;
        elements.errorMessage.style.display = 'block';
    }
}

function clearMessages() {
    if (elements.errorMessage) {
        elements.errorMessage.textContent = '';
        elements.errorMessage.style.display = 'none';
    }
}

// ============================================================================
// ТАБЫ: Проверка / Титульный лист
// ============================================================================

function showTab(tab) {
    const checkTab = document.getElementById('checkTab');
    const titleTab = document.getElementById('titleTab');
    const btnCheck = document.getElementById('tabBtnCheck');
    const btnTitle = document.getElementById('tabBtnTitle');

    if (tab === 'check') {
        checkTab.style.display = '';
        titleTab.style.display = 'none';
        btnCheck.classList.add('active');
        btnTitle.classList.remove('active');
    } else {
        checkTab.style.display = 'none';
        titleTab.style.display = '';
        btnCheck.classList.remove('active');
        btnTitle.classList.add('active');
    }
}

// ============================================================================
// ТИТУЛЬНЫЙ ЛИСТ — установка года по умолчанию
// ============================================================================

document.addEventListener('DOMContentLoaded', () => {
    const yearInput = document.getElementById('tp_year');
    if (yearInput && !yearInput.value) {
        yearInput.value = new Date().getFullYear();
    }
});

// ============================================================================
// ТИТУЛЬНЫЙ ЛИСТ — генерация
// ============================================================================

async function generateTitlePage() {
    const errorEl = document.getElementById('titleError');
    const successEl = document.getElementById('titleSuccess');
    const btn = document.getElementById('generateTitleBtn');

    // Скрываем сообщения
    errorEl.style.display = 'none';
    successEl.style.display = 'none';

    // Собираем данные
    const formData = new FormData();
    formData.append('work_title',     document.getElementById('tp_work_title').value);
    formData.append('work_number',    document.getElementById('tp_work_number').value);
    formData.append('specialty_code', document.getElementById('tp_specialty_code').value);
    formData.append('specialty_name', document.getElementById('tp_specialty_name').value);
    formData.append('subject',        document.getElementById('tp_subject').value);
    formData.append('group',          document.getElementById('tp_group').value);
    formData.append('student_id',     document.getElementById('tp_student_id').value);
    formData.append('student_name',   document.getElementById('tp_student_name').value);
    formData.append('teacher_name',   document.getElementById('tp_teacher_name').value);
    formData.append('city',           document.getElementById('tp_city').value);
    formData.append('year',           document.getElementById('tp_year').value);

    // Рамка
    const templateInput = document.getElementById('tp_template');
    if (templateInput && templateInput.files.length > 0) {
        formData.append('template', templateInput.files[0]);
    }

    // Валидация
    const studentName = document.getElementById('tp_student_name').value.trim();
    if (!studentName) {
        errorEl.textContent = '⚠️ Укажите ФИО студента';
        errorEl.style.display = 'block';
        return;
    }

    // Блокируем кнопку
    btn.disabled = true;
    btn.textContent = '⏳ Генерация...';

    try {
        const response = await fetch('/api/title-page/', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            // Ошибка — читаем JSON
            const contentType = response.headers.get('content-type') || '';
            if (contentType.includes('json')) {
                const data = await response.json();
                throw new Error(data.error || 'Ошибка сервера');
            }
            throw new Error(`HTTP ${response.status}`);
        }

        // Успех — скачиваем файл
        const blob = await response.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;

        // Имя файла из Content-Disposition или дефолтное
        const disposition = response.headers.get('Content-Disposition') || '';
        const match = disposition.match(/filename="?(.+?)"?$/);
        a.download = match ? match[1] : 'Титульный_лист.docx';

        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);

        successEl.textContent = '✅ Титульный лист сгенерирован и скачан!';
        successEl.style.display = 'block';

        console.log('✅ Титульный лист скачан');

    } catch (error) {
        console.error('❌ Ошибка генерации:', error);
        errorEl.textContent = '⚠️ ' + error.message;
        errorEl.style.display = 'block';
    } finally {
        btn.disabled = false;
        btn.textContent = '📄 Сгенерировать титульный лист';
    }
}