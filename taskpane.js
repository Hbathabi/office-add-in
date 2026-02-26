// Global state management
let documentClassification = {
    isClassified: false,
    classification: null,
    documentId: null,
    isInitialized: false
};

// Updated classification levels configuration
const CLASSIFICATION_LEVELS = {
    public: { 
        label: 'Public', 
        color: '#107c10', 
        order: 1,
        hasSublevels: false
    },
    confidential_internal: { 
        label: 'Confidential - Internal User', 
        color: '#ff8c00', 
        order: 2,
        parent: 'confidential',
        sublevel: 'internal'
    },
    confidential_external: { 
        label: 'Confidential - External User', 
        color: '#d83b01', 
        order: 3,
        parent: 'confidential',
        sublevel: 'external'
    },
    secret_internal: { 
        label: 'Secret - Internal User', 
        color: '#a4262c', 
        order: 4,
        parent: 'secret',
        sublevel: 'internal'
    },
    secret_external: { 
        label: 'Secret - External User', 
        color: '#8b0000', 
        order: 5,
        parent: 'secret',
        sublevel: 'external'
    },
    topsecret_internal: { 
        label: 'Top Secret - Internal User', 
        color: '#4b0000', 
        order: 6,
        parent: 'topsecret',
        sublevel: 'internal'
    },
    topsecret_external: { 
        label: 'Top Secret - External User', 
        color: '#2b0000', 
        order: 7,
        parent: 'topsecret',
        sublevel: 'external'
    }
};

// Classification hierarchy for dropdown
const CLASSIFICATION_HIERARCHY = {
    public: {
        label: 'Public',
        value: 'public',
        sublevels: null
    },
    confidential: {
        label: 'Confidential',
        value: null,
        sublevels: {
            internal: { label: 'Internal User', value: 'confidential_internal' },
            external: { label: 'External User', value: 'confidential_external' }
        }
    },
    secret: {
        label: 'Secret',
        value: null,
        sublevels: {
            internal: { label: 'Internal User', value: 'secret_internal' },
            external: { label: 'External User', value: 'secret_external' }
        }
    },
    topsecret: {
        label: 'Top Secret',
        value: null,
        sublevels: {
            internal: { label: 'Internal User', value: 'topsecret_internal' },
            external: { label: 'External User', value: 'topsecret_external' }
        }
    }
};

// Event handlers and state
let saveAttemptBlocked = false;
let modalResolver = null;

// Show notification function - DEFINE THIS FIRST
function showNotification(message, type = 'info') {
    console.log(`${type.toUpperCase()}: ${message}`);
    

    // Create notification element
    const notification = document.createElement('div');
    notification.className = `notification notification-${type}`;
    notification.textContent = message;
    
    // Style notification
    Object.assign(notification.style, {
        position: 'fixed',
        top: '20px',
        right: '20px',
        padding: '12px 20px',
        borderRadius: '4px',
        color: 'white',
        fontWeight: '500',
        zIndex: '1001',
        maxWidth: '300px',
        boxShadow: '0 4px 12px rgba(0, 0, 0, 0.3)',
        animation: 'slideInRight 0.3s ease-out'
    });
    
    // Set background color based on type
    const colors = {
        success: '#107c10',
        error: '#d13438',
        warning: '#ff8c00',
        info: '#0078d4'
    };
    notification.style.backgroundColor = colors[type] || colors.info;
    
    // Add to document
    document.body.appendChild(notification);
    
    // Remove after delay
    setTimeout(() => {
        notification.style.animation = 'slideOutRight 0.3s ease-in';
        setTimeout(() => {
            if (notification.parentNode) {
                notification.parentNode.removeChild(notification);
            }
        }, 300);
    }, 3000);
}

// Office initialization
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Word) {
        console.log('Office add-in initialized');
        try {
            await initializeClassificationSystem();
            setupEventHandlers();
            setupKeyboardInterceptors();
        } catch (error) {
            console.error('Error during initialization:', error);
            showNotification('Failed to initialize add-in', 'error');
        }
    }
});

// Main initialization function
async function initializeClassificationSystem() {
    try {
        console.log('Initializing classification system...');
        
        // Check existing classification
        const hasClassification = await checkDocumentClassification();
        console.log('Has existing classification:', hasClassification);
        
        // Update UI
        updateClassificationDisplay();
        
        // Mark as initialized
        documentClassification.isInitialized = true;
        
        // Prompt for classification if document is new and unclassified
        if (!hasClassification) {
            setTimeout(async () => {
                try {
                    await promptForClassificationIfNeeded('opening this document');
                } catch (error) {
                    console.log('Classification prompt was cancelled or failed:', error.message);
                }
            }, 2000); // Delay to ensure UI is ready
        }
        
    } catch (error) {
        console.error('Error initializing classification system:', error);
        showNotification('Error initializing classification system', 'error');
    }
}

// Check if document has existing classification
async function checkDocumentClassification() {
    try {
        return await Word.run(async (context) => {
            const properties = context.document.properties.customProperties;
            properties.load("items");
            
            await context.sync();
            
            const classificationProperty = properties.items.find(
                prop => prop.key === "DocumentClassification"
            );
            
            if (classificationProperty) {
                documentClassification.isClassified = true;
                documentClassification.classification = classificationProperty.value;
                console.log('Found existing classification:', classificationProperty.value);
                return true;
            }
            
            documentClassification.isClassified = false;
            documentClassification.classification = null;
            return false;
        });
    } catch (error) {
        console.error('Error checking document classification:', error);
        showNotification('Error checking document classification', 'error');
        return false;
    }
}

// Set document classification
async function setDocumentClassification(classification) {
    try {
        return await Word.run(async (context) => {
            // Set custom property
            const properties = context.document.properties.customProperties;
            properties.load("items");
            await context.sync();
            
            // Remove existing classification if present
            const existing = properties.items.find(
                prop => prop.key === "DocumentClassification"
            );
            if (existing) {
                existing.delete();
            }
            
            // Add new classification
            properties.add("DocumentClassification", classification);
            
            // Add classification header
            await addClassificationHeader(context, classification);
            
            await context.sync();
            
            // Update local state
            documentClassification.is
Classified = true;
            documentClassification.classification = classification;
            
            console.log('Classification set to:', classification);
            return true;
        });
    } catch (error) {
        console.error('Error setting document classification:', error);
        showNotification('Error setting document classification', 'error');
        throw error;
    }
}

// Add classification header to document - SIMPLE CENTERED VERSION
async function addClassificationHeader(context, classification) {
    try {
        const sections = context.document.sections;
        sections.load("items");
        await context.sync();
        
        if (sections.items.length > 0) {
            // FIXED: Use [0] instead of [2]
            const header = sections.items[0].getHeader("Primary");
            header.clear();
            
            const classificationConfig = CLASSIFICATION_LEVELS[classification];
            const classificationText = `${classificationConfig?.label || classification.toUpperCase()}`;
            
            // Simple approach - just insert paragraph and center it
            const paragraph = header.insertParagraph(classificationText, Word.InsertLocation.start);
            
            // Style the text
            paragraph.font.bold = true;
            paragraph.font.size = 10;
            paragraph.font.color = classificationConfig?.color || '#000000';
            
            // Center the paragraph
            paragraph.alignment = Word.Alignment.centered;
            
            // Add some spacing
            paragraph.spaceAfter = 12;
            paragraph.spaceBefore = 6;
            
            await context.sync();
            console.log('Classification header added and centered');
        }
    } catch (error) {
        console.error('Error adding classification header:', error);
        showNotification('Error adding classification header', 'warning');
    }
}

// Remove document classification
async function removeDocumentClassification() {
    try {
        return await Word.run(async (context) => {
            // Remove custom property
            const properties = context.document.properties.customProperties;
            properties.load("items");
            await context.sync();
            
            const existing = properties.items.find(
                prop => prop.key === "DocumentClassification"
            );
            if (existing) {
                existing.delete();
            }
            
            // Clear header
            const sections = context.document.sections;
            sections.load("items");
            await context.sync();
            
            if (sections.items.length > 0) {
                // FIXED: Use [0] instead of [2]
                const header = sections.items[0].getHeader("Primary");
                header.clear();
            }
            
            await context.sync();
            
            // Update local state
            documentClassification.isClassified = false;
            documentClassification.classification = null;
            
            console.log('Classification removed');
            return true;
        });
    } catch (error) {
        console.error('Error removing document classification:', error);
        showNotification('Error removing document classification', 'error');
        throw error;
    }
}

// Handle primary classification dropdown change
function handlePrimaryClassificationChange(primaryValue, sublevelContainer, sublevelDropdown, preview, previewText, setButton) {
    if (primaryValue === 'public') {
        // Public doesn't need sublevel
        sublevelContainer.style.display = 'none';
        sublevelDropdown.value = '';
        updatePreview('public', null, preview, previewText);
        setButton.disabled = false;
    } else if (primaryValue && CLASSIFICATION_HIERARCHY[primaryValue]) {
        // Show sublevel dropdown for other classifications
        sublevelContainer.style.display = 'block';
        sublevelDropdown.value = '';
        preview.style.display = 'none';
        setButton.disabled = true;
    } else {
        // No selection
        sublevelContainer.style.display = 'none';
        sublevelDropdown.value = '';
        preview.style.display = 'none';
        setButton.disabled = true;
    }
}

// Handle sublevel dropdown change
function handleSublevelChange(primaryValue, sublevelValue, preview, previewText, setButton) {
    if (primaryValue && sublevelValue) {
        const classificationKey = `${primaryValue}_${sublevelValue}`;
        updatePreview(classificationKey, sublevelValue, preview, previewText);
        setButton.disabled = false;
    } else {
        preview.style.display = 'none';
        setButton.disabled = true;
    }
}

// Update preview display
function updatePreview(classificationKey, sublevel, preview, previewText) {
    const config = CLASSIFICATION_LEVELS[classificationKey];
    if (config) {
        previewText.textContent = config.label;
        previewText.style.color = config.color;
        preview.style.display = 'block';
    }
}

// Getselected classification from dropdowns
function getSelectedClassification() {
    const primaryDropdown = document.getElementById('classification-primary');
    const sublevelDropdown = document.getElementById('classification-sublevel');
    
    if (!primaryDropdown) return null;
    
    const primaryValue = primaryDropdown.value;
    
    if (primaryValue === 'public') {
        return 'public';
    } else if (primaryValue && sublevelDropdown && sublevelDropdown.value) {
        return `${primaryValue}_${sublevelDropdown.value}`;
    }
    
    return null;
}

// Reset dropdowns to initial state
function resetDropdowns() {
    const primaryDropdown = document.getElementById('classification-primary');
    const sublevelDropdown = document.getElementById('classification-sublevel');
    const sublevelContainer = document.getElementById('sublevel-container');
    const preview = document.getElementById('classification-preview');
    const setButton = document.getElementById('set-classification');
    
    if (primaryDropdown) primaryDropdown.value = '';
    if (sublevelDropdown) sublevelDropdown.value = '';
    if (sublevelContainer) sublevelContainer.style.display = 'none';
    if (preview) preview.style.display = 'none';
    if (setButton) setButton.disabled = true;
}

// Get selected classification from modal dropdowns
function getModalSelectedClassification() {
    const modalPrimaryDropdown = document.getElementById('modal-classification-primary');
    const modalSublevelDropdown = document.getElementById('modal-classification-sublevel');
    
    if (!modalPrimaryDropdown) return null;
    
    const primaryValue = modalPrimaryDropdown.value;
    
    if (primaryValue === 'public') {
        return 'public';
    } else if (primaryValue && modalSublevelDropdown && modalSublevelDropdown.value) {
        return `${primaryValue}_${modalSublevelDropdown.value}`;
    }
    
    return null;
}

// Reset modal dropdowns
function resetModalDropdowns() {
    const modalPrimaryDropdown = document.getElementById('modal-classification-primary');
    const modalSublevelDropdown = document.getElementById('modal-classification-sublevel');
    const modalSublevelContainer = document.getElementById('modal-sublevel-container');
    const modalPreview = document.getElementById('modal-classification-preview');
    const modalConfirm = document.getElementById('modal-confirm');
    
    if (modalPrimaryDropdown) modalPrimaryDropdown.value = '';
    if (modalSublevelDropdown) modalSublevelDropdown.value = '';
    if (modalSublevelContainer) modalSublevelContainer.style.display = 'none';
    if (modalPreview) modalPreview.style.display = 'none'; // FIXED: removed typo
    if (modalConfirm) modalConfirm.disabled = true;
}

// Setup event handlers for UI elements
function setupEventHandlers() {
    try {
        // Primary classification dropdown
        const primaryDropdown = document.getElementById('classification-primary');
        const sublevelContainer = document.getElementById('sublevel-container');
        const sublevelDropdown = document.getElementById('classification-sublevel');
        const preview = document.getElementById('classification-preview');
        const previewText = document.getElementById('preview-text');
        const setButton = document.getElementById('set-classification');

        if (primaryDropdown) {
            primaryDropdown.addEventListener('change', () => {
                handlePrimaryClassificationChange(
                    primaryDropdown.value,
                    sublevelContainer,
                    sublevelDropdown,
                    preview,
                    previewText,
                    setButton
                );
            });
        }

        if (sublevelDropdown) {
            sublevelDropdown.addEventListener('change', () => {
                handleSublevelChange(
                    primaryDropdown.value,
                    sublevelDropdown.value,
                    preview,
                    previewText,
                    setButton
                );
            });
        }
        
        // Set classification button
        if (setButton) {
            setButton.addEventListener('click', async () => {
                const selectedClassification = getSelectedClassification();
                if (selectedClassification) {
                    try {
                        await setDocumentClassification(selectedClassification);
                        updateClassificationDisplay();
                        showNotification('Classification set successfully', 'success');
                        resetDropdowns();
                    } catch (error) {
                        console.error('Error setting classification:', error);
                        showNotification('Error setting classification', 'error');
                    }
                }
            });
        }
        
        // Remove classification button
        const removeButton = document.getElementById('remove-classification');
        if (removeButton) {
            removeButton.addEventListener('click', async () => {
                if (confirm('Are you sure you want to remove the document classification?')) {
                    try {
                        await removeDocumentClassification();
                        updateClassificationDisplay();
                        showNotification('Classification removed', 'success');
                        resetDropdowns();
                    } catch (error) {
                        console.error('Error removing classification:', error);
                        showNotification('Error removing classification', 'error');
                    }
                }
            });
        }
        
        // Save document button
        const saveButton = document.getElementById('save-document');
        if (saveButton) {
            saveButton.addEventListener('click', async () => {
                await handleSaveDocument();
            });
        }
        
        // Refresh status button
        const refreshButton = document.getElementById('refresh-status');
        if (refreshButton) {
            refreshButton.addEventListener('click', async () => {
                try {
                    await checkDocumentClassification();
                    updateClassificationDisplay();
                    showNotification('Status refreshed', 'info');
                } catch (error) {
                    console.error('Error refreshing status:', error);
                    showNotification('Error refreshing status', 'error');
                }
            });
        }
        
        // Modal event handlers
        setupModalEventHandlers();
        
    } catch (error) {
        console.error('Error setting up event handlers:', error);
        showNotification('Error setting up event handlers', 'error');
    }
}

// Setup modal event handlers with dropdown support
function setupModalEventHandlers() {
    try {
        const modal = document.getElementById('classification-modal');
        const modalClose
 = document.getElementById('modal-close');
        const modalCancel = document.getElementById('modal-cancel');
        const modalConfirm = document.getElementById('modal-confirm');
        
        // Modal dropdown handlers
        const modalPrimaryDropdown = document.getElementById('modal-classification-primary');
        const modalSublevelContainer = document.getElementById('modal-sublevel-container');
        const modalSublevelDropdown = document.getElementById('modal-classification-sublevel');
        const modalPreview = document.getElementById('modal-classification-preview');
        const modalPreviewText = document.getElementById('modal-preview-text');
        
        if (!modal || !modalClose || !modalCancel || !modalConfirm) {
            console.warn('Modal elements not found, skipping modal setup');
            return;
        }
        
        // Modal primary dropdown change
        if (modalPrimaryDropdown) {
            modalPrimaryDropdown.addEventListener('change', () => {
                handlePrimaryClassificationChange(
                    modalPrimaryDropdown.value,
                    modalSublevelContainer,
                    modalSublevelDropdown,
                    modalPreview,
                    modalPreviewText,
                    modalConfirm
                );
            });
        }
        
        // Modal sublevel dropdown change
        if (modalSublevelDropdown) {
            modalSublevelDropdown.addEventListener('change', () => {
                handleSublevelChange(
                    modalPrimaryDropdown.value,
                    modalSublevelDropdown.value,
                    modalPreview,
                    modalPreviewText,
                    modalConfirm
                );
            });
        }
        
        // Close modal handlers
        modalClose.addEventListener('click', () => closeModal(false));
        modalCancel.addEventListener('click', () => closeModal(false));
        
        // Confirm modal handler
        modalConfirm.addEventListener('click', async () => {
            const selectedClassification = getModalSelectedClassification();
            if (selectedClassification) {
                try {
                    await setDocumentClassification(selectedClassification);
                    updateClassificationDisplay();
                    closeModal(true);
                    showNotification('Classification set successfully', 'success');
                } catch (error) {
                    console.error('Error setting classification from modal:', error);
                    showNotification('Error setting classification', 'error');
                    closeModal(false);
                }
            }
        });
        
        // Close modal when clicking outside
        modal.addEventListener('click', (event) => {
            if (event.target === modal) {
                closeModal(false);
            }
        });
        
    } catch (error) {
        console.error('Error setting up modal handlers:', error);
        showNotification('Error setting up modal handlers', 'error');
    }
}

// Setup keyboard interceptors
function setupKeyboardInterceptors() {
    try {
        // Intercept Ctrl+S (Save)
        document.addEventListener('keydown', async (event) => {
            if ((event.ctrlKey || event.metaKey) && event.key === 's') {
                event.preventDefault(); // FIXED: joined the line
                await handleSaveDocument();
            }
        });
        
        // Intercept browser close/refresh
        window.addEventListener('beforeunload', (event) => {
            if (!documentClassification.isClassified && documentClassification.isInitialized) {
                event.preventDefault();
                event.returnValue = 'Document must be classified before closing.';
                return 'Document must be classified before closing.';
            }
        });
        
    } catch (error) {
        console.error('Error setting up keyboard interceptors:', error);
        showNotification('Error setting up keyboard interceptors', 'error');
    }
}

// Handle save document with classification check
async function handleSaveDocument() {
    try {
        if (!documentClassification.isClassified) {
            const classified = await promptForClassificationIfNeeded('saving this document');
            if (!classified) {
                showNotification('Save cancelled - document must be classified', 'warning');
                return false;
            }
        }
        
        // Proceed with save
        await Word.run(async (context) => {
            await context.document.save();
            await context.sync();
        });
        
        showNotification('Document saved successfully', 'success');
        return true;
        
    } catch (error) {
        console.error('Error saving document:', error);
        showNotification('Error saving document', 'error');
        return false;
    }
}

// Prompt for classification if needed
async function promptForClassificationIfNeeded(action = 'proceeding') {
    if (documentClassification.isClassified) {
        return true;
    }
    
    try {
        const result = await showClassificationModal(action);
        return result;
    } catch (error) {
        console.error('Classification prompt cancelled or failed:', error);
        return false;
    }
}

// Show classification modal
function showClassificationModal(action = 'proceeding') {
    return new Promise((resolve, reject) => {
        const modal = document.getElementById('classification-modal');
        const modalMessage = document.getElementById('modal-message');
        
        if (!modal || !modalMessage) {
            console.error('Modal elements not found');
            reject(new Error('Modal elements not found'));
            return;
        }
        
        // Reset modal state
        resetModalDropdowns();
        
        // Set message
        modalMessage.textContent = `Please select a classification level before ${action}:`;
        
        // Show modal
        modal.style.display = 'flex';
        
        // Store resolver for modal
 handlers
        modalResolver = { resolve, reject };
    });
}

// Close modal
function closeModal(success) {
    const modal = document.getElementById('classification-modal');
    if (modal) {
        modal.style.display = 'none';
    }
    
    if (modalResolver) {
        if (success) {
            modalResolver.resolve(true);
        } else {
            modalResolver.reject(new Error('Modal cancelled'));
        }
        modalResolver = null;
    }
}

// Update classification display in UI
function updateClassificationDisplay() {
    try {
        const classificationText = document.getElementById('classification-text');
        const classificationLevel = document.getElementById('classification-level');
        
        if (!classificationText || !classificationLevel) {
            console.warn('Classification display elements not found');
            return;
        }
        
        if (documentClassification.isClassified && documentClassification.classification) {
            const level = documentClassification.classification;
            const config = CLASSIFICATION_LEVELS[level];
            
            classificationText.textContent = `Document is classified as: ${config?.label || level}`;
            classificationLevel.textContent = (config?.label || level).toUpperCase();
            classificationLevel.className = `classification-badge ${level.replace('_', '-')}`;
        } else {
            classificationText.textContent = 'Document is not classified';
            classificationLevel.textContent = 'UNCLASSIFIED';
            classificationLevel.className = 'classification-badge unclassified';
        }
        
        // Update remove button state
        const removeButton = document.getElementById('remove-classification');
        if (removeButton) {
            removeButton.disabled = !documentClassification.isClassified;
        }
        
    } catch (error) {
        console.error('Error updating classification display:', error);
        showNotification('Error updating display', 'error');
    }
}

// Add CSS animations for notifications
const style = document.createElement('style');
style.textContent = `
    @keyframes slideInRight {
        from {
            opacity: 0;
            transform: translateX(100%);
        }
        to {
            opacity: 1;
            transform: translateX(0);
        }
    }
    
    @keyframes slideOutRight {
        from {
            opacity: 1;
            transform: translateX(0);
        }
        to {
            opacity: 0;
            transform: translateX(100%);
        }
    }
`;
document.head.appendChild(style);

// Export functions for ribbon commands
window.classificationAddin = {
    setDocumentClassification,
    removeDocumentClassification,
    checkDocumentClassification,
    handleSaveDocument,
    promptForClassificationIfNeeded,
    showNotification
};

// Global error handler
window.addEventListener('error', (event) => {
    console.error('Global error:', event.error);
    if (typeof showNotification === 'function') {
        showNotification('An unexpected error occurred', 'error');
    }
});

// Global unhandled promise rejection handler
window.addEventListener('unhandledrejection', (event) => {
    console.error('Unhandled promise rejection:', event.reason);
    if (typeof showNotification === 'function') {
        showNotification('An unexpected error occurred', 'error');
    }
});