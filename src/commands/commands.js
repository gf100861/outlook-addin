/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(() => {
    // If needed, Office.js is ready to be called
});

/**
 * Handles the OnMessageSend event.
 * @param {Office.AddinCommands.Event} event The event object.
 */
async function onMessageSendHandler(event) {
    console.log("🚀 onMessageSendHandler function started! (Final Version)");

    try {
        // --- Keyword Lists ---
        const imageKeywords = [
            "image", "images", "picture", "photo", "screenshot",
            "图片", "照片", "截图"
        ];
        
        const generalKeywords = [
            "attach", "attached", "attaching", "attachment", "attachments",
            "enclosed", "file", "files", "find attached", "see attached", 
            "review the attached", "including", "附件", "附上", "见附件", 
            "查收", "请查收", "文件", "文档", "报告", "简历", "表格", "演示",
            ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".zip", ".rar"
        ];
        
        const item = Office.context.mailbox.item;

        // --- Get Subject and Body ---
        const subject = await new Promise((resolve) => {
            item.subject.getAsync((result) => resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : ""));
        });
        const lowerCaseSubject = subject.toLowerCase();

        const body = await new Promise((resolve) => {
            item.body.getAsync(Office.CoercionType.Text, (result) => resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : ""));
        });
        const lowerCaseBody = body.toLowerCase();
        const fullText = lowerCaseSubject + " " + lowerCaseBody;

        // --- Step 1: Keyword Detection ---
        const imageKeywordFound = imageKeywords.some(keyword => fullText.includes(keyword));
        // Ensure general keyword check doesn't overlap with image keywords if an image keyword is already found
        const generalKeywordFound = !imageKeywordFound && generalKeywords.some(keyword => fullText.includes(keyword));

        console.log(`🖼️ Image keyword found? ${imageKeywordFound}`);
        console.log(`📎 General keyword found? ${generalKeywordFound}`);

        // If no keywords are found at all, allow sending immediately.
        if (!imageKeywordFound && !generalKeywordFound) {
            console.log("➡️ No keywords found. Allowing send.");
            event.completed({ allowEvent: true });
            return;
        }

        // --- Step 2: Attachment Validation (only if a keyword was found) ---
        console.log("📎 Keyword detected. Validating attachments...");
        const attachments = await new Promise((resolve) => {
            item.getAttachmentsAsync((result) => resolve(result.status === Office.AsyncResultStatus.Succeeded ? result.value : []));
        });

        let allowSend = false;

        if (generalKeywordFound) {
            // General keywords require a REAL (non-inline) attachment.
            console.log("General keyword found. Checking for REAL attachments...");
            if (attachments.some(att => !att.isInline)) {
                console.log("✅ Real attachment found.");
                allowSend = true;
            } else {
                console.log("❌ No real attachment found.");
            }
        } else if (imageKeywordFound) {
            // Image keywords allow ANY attachment (inline or real).
            console.log("Image keyword found. Checking for ANY attachments...");
            if (attachments.length > 0) {
                console.log("✅ An attachment (inline or real) was found.");
                allowSend = true;
            } else {
                console.log("❌ No attachments of any kind were found.");
            }
        }

        // --- Step 3: Final Decision ---
        if (allowSend) {
            console.log("✅ Requirements met. Allowing send.");
            event.completed({ allowEvent: true });
        } else {
            console.warn("❗ Requirements NOT met. Blocking send.");
            event.completed({
                allowEvent: false,
                errorMessage: "您似乎忘记添加附件了。(You seem to have forgotten an attachment.)",
            });
        }

    } catch (error) {
        console.error("❌ Unexpected error occurred:", error);
        event.completed({ allowEvent: true });
    }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);