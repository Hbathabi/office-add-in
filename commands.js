// Commands for ribbon buttons
Office.onReady(() => {
    console.log('Commands initialized');
});

// Save with classification check
async function saveWithClassification(event) {
    try {
        // Check if document is classified
        const isClassified = await checkDocumentClassificationStatus();
        
        if (!isClassified) {
            // Show dialog to classify document
            Office.context.ui.displayDialogAsync(
                'https://localhost:3000/taskpane.html',
                { height: 60, width: 40 },
                (result) => {
                    if (result.status
 === Office.AsyncResultStatus.Succeeded) {
                        const dialog = result.value;
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
                            if (arg.message === 'classified') {
                                performSave();
                            }
                            dialog.close();
                        });
                    }
                }
            );
        } else {
            await performSave();
        }
        
        event.completed();
    } catch (error) {
        console.error('Error in saveWithClassification:', error);
        event.completed();
    }
}

// Check classification status
async function checkDocumentClassificationStatus() {
    try {
        return await Word.run(async (context) => {
            const properties = context.document.properties.customProperties;
            properties.load("items");
            await context.sync();
            
            const classificationProperty = properties.items.find(
                prop => prop.key === "DocumentClassification"
            );
            
            return !!classificationProperty;
        });
    } catch (error) {
        console.error('Error checking classification status:', error);
        return false;
    }
}

// Perform actual save
async function performSave() {
    try {
        await Word.run(async (context) => {
            await context.document.save();
            await context.sync();
        });
        
        // Show success notification
        Office.context.ui.displayDialogAsync(
            'data:text/html,<html><body><script>window.parent.postMessage("Document saved successfully", "*");</script></body></html>',
            { height: 20, width: 30 }
        );
    } catch (error) {
        console.error('Error saving document:', error);
    }
}

// Register functions for ribbon commands
Office.actions.associate("saveWithClassification", saveWithClassification);