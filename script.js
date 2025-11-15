const tool = {
    currentIndex: 0,
    conversations: [],
    annotations: {},
    buckets: [
        "Bot Response",
        "HVA",
        "AB feature/HVA Related Query",
        "Personalized/Account-Specific Queries",
        "Promo & Freebie Related Queries",
        "Help-page/Direct Customer Service",
        "BP for Non-Profit Organisation Related Query",
        "Personal Prime Related Query",
        "Customer Behavior",
        "Other Queries",
        "Overall Observations"
    ],
    hvaOptions: [
        "3WM (3-Way Match)",
        "Account Authority",
        "Add User",
        "Analytics",
        "ATEP (Amazon Tax Exemption Program)",
        "Business Lists",
        "Business Order Information",
        "Custom Quotes",
        "Guided Buying",
        "PBI",
        "Quantity Discount",
        "Shared Settings",
        "SSO",
        "Subscibe & Save (formerly Recurring Delivery)"
    ]
};

const elements = {
    uploadScreen: document.getElementById('upload-screen'),
    mainInterface: document.getElementById('main-interface'),
    uploadBox: document.getElementById('upload-box'),
    fileInput: document.getElementById('excel-upload'),
    uploadStatus: document.getElementById('upload-status'),
    conversationDisplay: document.getElementById('conversation-display'),
    conversationInfo: document.getElementById('conversation-info'),
    bucketArea: document.getElementById('bucket-area'),
    prevBtn: document.getElementById('prev-btn'),
    nextBtn: document.getElementById('next-btn'),
    saveBtn: document.getElementById('save-btn'),
    downloadBtn: document.getElementById('download-btn'),
    progress: document.getElementById('progress'),
    progressText: document.getElementById('progress-text'),
    statusMessage: document.getElementById('status-message'),
    loadingSpinner: document.getElementById('loading-spinner')
};

function createBucketUI() {
    tool.buckets.forEach(bucket => {
        let contentHTML = '';
        
        if (bucket === "Bot Response") {
            // Radio buttons for Bot Response
            contentHTML = `
                <div class="radio-group">
                    <label class="radio-option">
                        <input type="radio" name="${bucket}-option" value="Accurate response">
                        <span>Accurate response</span>
                    </label>
                    <label class="radio-option">
                        <input type="radio" name="${bucket}-option" value="Inaccurate response">
                        <span>Inaccurate response</span>
                    </label>
                </div>
            `;
        } else if (bucket === "HVA") {
            contentHTML = `
                <select class="form-select hva-dropdown" name="${bucket}-select">
                    <option value="">Select HVA type...</option>
                    ${tool.hvaOptions.map(option => `<option value="${option}">${option}</option>`).join('')}
                </select>
            `;
        } else {
            contentHTML = `<textarea placeholder="Add comments for ${bucket}" name="${bucket}" rows="3"></textarea>`;
        }
        
        const bucketHTML = `
            <div class="bucket" data-bucket="${bucket}">
                <label class="bucket-label">
                    <input type="checkbox" name="${bucket}">
                    <span>${bucket}</span>
                </label>
                <div class="bucket-comment">${contentHTML}</div>
            </div>
        `;
        elements.bucketArea.insertAdjacentHTML('beforeend', bucketHTML);
    });
}

// File Upload Handlers
elements.uploadBox.addEventListener('click', () => {
    elements.fileInput.click();
});

elements.fileInput.addEventListener('change', async (event) => {
    const file = event.target.files[0];
    if (!file) {
        showStatus('⚠️ No file selected', 'warning');
        return;
    }
    if (!file.name.endsWith('.xlsx')) {
        showStatus('❌ Please select an Excel (.xlsx) file', 'error');
        return;
    }
    try {
        showLoading(true);
        showStatus('📂 Loading file...', 'info');
        const data = await readExcelFile(file);
        if (!data || data.length === 0) {
            throw new Error('No data found in file');
        }
        processExcelData(data);
        elements.uploadScreen.style.display = 'none';
        elements.mainInterface.style.display = 'flex';
        showStatus('✅ File loaded successfully!', 'success');
    } catch (error) {
        showStatus('❌ Error: ' + (error.message || 'Failed to load file'), 'error');
    } finally {
        showLoading(false);
    }
});

// Drag and drop handlers
elements.uploadBox.addEventListener('dragover', (e) => {
    e.preventDefault();
    elements.uploadBox.classList.add('dragover');
});

elements.uploadBox.addEventListener('dragleave', () => {
    elements.uploadBox.classList.remove('dragover');
});

elements.uploadBox.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.uploadBox.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file && file.name.endsWith('.xlsx')) {
        elements.fileInput.files = e.dataTransfer.files;
        elements.fileInput.dispatchEvent(new Event('change'));
    } else {
        showStatus('❌ Please select an Excel (.xlsx) file', 'error');
    }
});

// File reading and processing functions
async function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, {type: 'array'});
                const sheetName = workbook.SheetNames[0];
                const sheet = workbook.Sheets[sheetName];
                const jsonData = XLSX.utils.sheet_to_json(sheet);
                resolve(jsonData);
            } catch (error) {
                reject(new Error('Failed to parse Excel file'));
            }
        };
        reader.onerror = () => reject(new Error('Failed to read file'));
        reader.readAsArrayBuffer(file);
    });
}

function processExcelData(rawData) {
    const groupedData = {};
    rawData.forEach(row => {
        if (!groupedData[row.Id]) {
            groupedData[row.Id] = [];
        }
        groupedData[row.Id].push(row);
    });
    tool.conversations = Object.values(groupedData);
    tool.currentIndex = 0;
    tool.annotations = {};
    updateProgressBar();
    displayConversation();
}

function displayConversation() {
    const conv = tool.conversations[tool.currentIndex];
    const lastMessage = conv[conv.length - 1];

    elements.conversationInfo.innerHTML = `
        <div class="info-item">
            <strong>ID:</strong> ${conv[0].Id}
        </div>
        <div class="info-item">
            <strong>Feedback:</strong> 
            <span class="badge ${lastMessage['Customer Feedback']?.toLowerCase() === 'negative' ? 'bg-danger' : 'bg-success'}">
                ${lastMessage['Customer Feedback'] || 'N/A'}
            </span>
        </div>
    `;

    let html = '<div class="messages">';
    conv.forEach(message => {
        if (message.llmGeneratedUserMessage) {
            html += `
                <div class="message customer">
                    <div class="message-header">👤 Customer</div>
                    ${message.llmGeneratedUserMessage}
                </div>
            `;
        }
        if (message.botMessage) {
            html += `
                <div class="message bot">
                    <div class="message-header">🤖 Bot</div>
                    ${message.botMessage}
                </div>
            `;
        }
    });
    html += '</div>';

    elements.conversationDisplay.innerHTML = html;
    updateProgressBar();
    loadAnnotations();
}

function updateProgressBar() {
    const progress = ((tool.currentIndex + 1) / tool.conversations.length) * 100;
    elements.progress.style.width = `${progress}%`;
    elements.progressText.textContent = `${tool.currentIndex + 1}/${tool.conversations.length} Conversations`;
}

function saveCurrentAnnotations() {
    const convId = tool.conversations[tool.currentIndex][0].Id;
    const hasAnnotations = tool.buckets.some(bucket =>
        document.querySelector(`input[name="${bucket}"]`).checked
    );

    if (!hasAnnotations) {
        showStatus('⚠️ Please select at least one bucket', 'warning');
        return;
    }

    tool.annotations[convId] = {};
    tool.buckets.forEach(bucket => {
        const checkbox = document.querySelector(`input[name="${bucket}"]`);
        if (checkbox.checked) {
            if (bucket === "Bot Response") {
                const selectedRadio = document.querySelector(`input[name="${bucket}-option"]:checked`);
                tool.annotations[convId][bucket] = selectedRadio ? selectedRadio.value : '';
            } else if (bucket === "HVA") {
                const dropdown = document.querySelector(`select[name="${bucket}-select"]`);
                tool.annotations[convId][bucket] = dropdown ? dropdown.value : '';
            } else {
                const textarea = document.querySelector(`textarea[name="${bucket}"]`);
                tool.annotations[convId][bucket] = textarea.value.trim();
            }
        }
    });

    showStatus('✅ Annotations saved!', 'success');
}

function loadAnnotations() {
    const convId = tool.conversations[tool.currentIndex][0].Id;
    const savedAnnotations = tool.annotations[convId] || {};

    tool.buckets.forEach(bucket => {
        const bucketDiv = document.querySelector(`[data-bucket="${bucket}"]`);
        const checkbox = bucketDiv.querySelector('input[type="checkbox"]');
        const commentDiv = bucketDiv.querySelector('.bucket-comment');
        
        checkbox.checked = false;
        commentDiv.classList.remove('open');
        bucketDiv.classList.remove('checked');
        
        if (bucket === "Bot Response") {
            const radios = bucketDiv.querySelectorAll('input[type="radio"]');
            radios.forEach(radio => radio.checked = false);
        } else if (bucket === "HVA") {
            const dropdown = bucketDiv.querySelector('select');
            if (dropdown) dropdown.value = '';
        } else {
            const textarea = bucketDiv.querySelector('textarea');
            if (textarea) textarea.value = '';
        }
    });

    Object.entries(savedAnnotations).forEach(([bucket, value]) => {
        const bucketDiv = document.querySelector(`[data-bucket="${bucket}"]`);
        if (bucketDiv) {
            const checkbox = bucketDiv.querySelector('input[type="checkbox"]');
            const commentDiv = bucketDiv.querySelector('.bucket-comment');
            
            checkbox.checked = true;
            commentDiv.classList.add('open');
            bucketDiv.classList.add('checked');
            
            if (bucket === "Bot Response") {
                const radio = bucketDiv.querySelector(`input[type="radio"][value="${value}"]`);
                if (radio) radio.checked = true;
            } else if (bucket === "HVA") {
                const dropdown = bucketDiv.querySelector('select');
                if (dropdown) dropdown.value = value;
            } else {
                const textarea = bucketDiv.querySelector('textarea');
                if (textarea) textarea.value = value;
            }
        }
    });
}

function showStatus(message, type) {
    elements.statusMessage.textContent = message;
    elements.statusMessage.className = `status-message alert alert-${type}`;
    elements.statusMessage.style.display = 'block';
    setTimeout(() => {
        elements.statusMessage.style.display = 'none';
    }, 3000);
}

function showLoading(show) {
    elements.loadingSpinner.style.display = show ? 'flex' : 'none';
}

// Bucket interaction handlers
elements.bucketArea.addEventListener('change', (e) => {
    if (e.target.type === 'checkbox') {
        const bucketDiv = e.target.closest('.bucket');
        const commentDiv = bucketDiv.querySelector('.bucket-comment');
        const input = bucketDiv.querySelector('textarea, select, input[type="radio"]');

        if (e.target.checked) {
            commentDiv.classList.add('open');
            bucketDiv.classList.add('checked');
            setTimeout(() => input?.focus(), 300);
        } else {
            commentDiv.classList.remove('open');
            bucketDiv.classList.remove('checked');
            
            // Clear values based on bucket type
            const bucket = bucketDiv.getAttribute('data-bucket');
            if (bucket === "Bot Response") {
                const radios = bucketDiv.querySelectorAll('input[type="radio"]');
                radios.forEach(radio => radio.checked = false);
            } else if (bucket === "HVA") {
                const select = bucketDiv.querySelector('select');
                if (select) select.value = '';
            } else {
                const textarea = bucketDiv.querySelector('textarea');
                if (textarea) textarea.value = '';
            }
        }
    }
});

// Excel download helper
function s2ab(s) {
    const buf = new ArrayBuffer(s.length);
    const view = new Uint8Array(buf);
    for (let i = 0; i < s.length; i++) {
        view[i] = s.charCodeAt(i) & 0xFF;
    }
    return buf;
}

// Download handler
elements.downloadBtn.addEventListener('click', async () => {
    try {
        if (Object.keys(tool.annotations).length === 0) {
            showStatus('⚠️ No annotations to download', 'warning');
            return;
        }

        showLoading(true);
        showStatus('💾 Preparing download...', 'info');
        
        const annotatedData = [];
        tool.conversations.forEach(conv => {
            const convId = conv[0].Id;
            const savedAnnotations = tool.annotations[convId];
            
            if (savedAnnotations && Object.keys(savedAnnotations).length > 0) {
                conv.forEach((message, index) => {
                    const isFirstMessage = index === 0;
                    const isLastMessage = index === conv.length - 1;
                    
                    const row = {
                        'Id': message.Id,
                        'llmGeneratedUserMessage': message.llmGeneratedUserMessage || '',
                        'botMessage': message.botMessage || '',
                        'Customer Feedback': isLastMessage ? message['Customer Feedback'] || '' : ''
                    };

                    if (isFirstMessage) {
                        tool.buckets.forEach(bucket => {
                            row[bucket] = savedAnnotations[bucket] || '';
                        });
                    } else {
                        tool.buckets.forEach(bucket => {
                            row[bucket] = '';
                        });
                    }
                    
                    annotatedData.push(row);
                });
            }
        });

        const ws = XLSX.utils.json_to_sheet(annotatedData);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Annotations");
        
        const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
        const blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
        const url = window.URL.createObjectURL(blob);
        
        const a = document.createElement('a');
        a.href = url;
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        a.download = `annotated_conversations_${timestamp}.xlsx`;
        document.body.appendChild(a);
        a.click();
        
        window.URL.revokeObjectURL(url);
        document.body.removeChild(a);

        const annotatedCount = new Set(annotatedData.map(row => row.Id)).size;
        showStatus(`✅ Downloaded ${annotatedCount} conversation(s)!`, 'success');
    } catch (error) {
        showStatus('❌ Error downloading file', 'error');
    } finally {
        showLoading(false);
    }
});

// Navigation handlers
elements.prevBtn.addEventListener('click', () => {
    if (tool.currentIndex > 0) {
        tool.currentIndex--;
        displayConversation();
    } else {
        showStatus('⚠️ This is the first conversation', 'warning');
    }
});

elements.nextBtn.addEventListener('click', () => {
    if (tool.currentIndex < tool.conversations.length - 1) {
        tool.currentIndex++;
        displayConversation();
    } else {
        showStatus('⚠️ This is the last conversation', 'warning');
    }
});

// Save button handler
elements.saveBtn.addEventListener('click', saveCurrentAnnotations);

// Keyboard navigation
document.addEventListener('keydown', (e) => {
    if (elements.mainInterface.style.display === 'none') return;
    if (e.key === 'ArrowLeft') {
        elements.prevBtn.click();
    } else if (e.key === 'ArrowRight') {
        elements.nextBtn.click();
    } else if
      } else if (e.key === 's' && (e.ctrlKey || e.metaKey)) {
        e.preventDefault();
        elements.saveBtn.click();
    }
});

// Initialize
createBucketUI();

// Window event handlers
window.addEventListener('resize', () => {
    if (elements.mainInterface.style.display !== 'none') {
        updateProgressBar();
    }
});

window.addEventListener('beforeunload', (e) => {
    if (Object.keys(tool.annotations).length > 0) {
        e.preventDefault();
        e.returnValue = '';
    }
});

console.log('Tool initialized and ready! 🚀');
