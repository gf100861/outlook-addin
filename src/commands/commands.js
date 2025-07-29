/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// Make sure Office is ready before registering the handler
Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Handles the OnMessageSend event.
 * @param {Office.AddinCommands.Event} event The event object.
 */
async function onMessageSendHandler(event) {
  console.log("🚀 onMessageSendHandler function started! (Fallback Mode)");

  try {
    // 回退到最简单的字符串数组，明确列出所有可能性
    const keywords = [
      // English
      "attach", "attached", "attaching", "attachment", "attachments",
      "enclosed",
      "image", "images", // 明确列出单数和复数
      "file", "files",   // 明确列出单数和复数
      "find attached", "see attached", "review the attached",
      "including",

      // Chinese (中文)
      "附件", "附上", "见附件", "查收", "请查收",
      "文件", "文档", "报告", "简历", "表格", "演示",
      "照片",

      // File Extensions
      ".pdf", ".doc", ".docx", ".xls", ".xlsx", ".ppt", ".pptx", ".zip", ".rar"
    ];

    const item = Office.context.mailbox.item;

    // --- SUBJECT CHECK ---
    const subject = await new Promise((resolve, reject) => {
      item.subject.getAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) reject(result.error);
        else resolve(result.value);
      });
    });

    // 检查时，将邮件内容和关键字都转换为小写
    const lowerCaseSubject = subject.toLowerCase();
    const subjectContainsKeyword = keywords.some(keyword => lowerCaseSubject.includes(keyword.toLowerCase()));
    console.log(`📌 Subject: "${lowerCaseSubject}"`);
    console.log(`✅ Subject contains keyword? ${subjectContainsKeyword}`);

    // --- BODY CHECK ---
    const body = await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve(result.value);
        else {
          console.error("❌ getAsync(body) failed:", result.error);
          resolve("");
        }
      });
    });

    const lowerCaseBody = body.toLowerCase();
    const bodyContainsKeyword = keywords.some(keyword => lowerCaseBody.includes(keyword.toLowerCase()));
    console.log(`✅ Body contains keyword? ${bodyContainsKeyword}`);
    
    // --- 其余代码保持不变 ---
    const keywordDetected = subjectContainsKeyword || bodyContainsKeyword;

    if (!keywordDetected) {
      console.log("➡️ No keywords found. Allowing send.");
      event.completed({ allowEvent: true });
      return;
    }

    console.log("📎 Keywords detected. Checking attachments...");
    const attachments = await new Promise((resolve, reject) => {
      item.getAttachmentsAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Failed) reject(result.error);
        else resolve(result.value);
      });
    });

    const hasRealAttachment = attachments.some(att => !att.isInline);

    if (hasRealAttachment) {
      console.log("✅ Real attachment exists. Allowing send.");
      event.completed({ allowEvent: true });
    } else {
      console.warn("❗ No real attachment. Blocking send.");
      event.completed({
        allowEvent: false,
        errorMessage: "您似乎忘记添加附件了。(You seem to have forgotten an attachment.)",
        cancelLabel: "添加附件 (Add Attachment)"
      });
    }

  } catch (error) {
    console.error("❌ Unexpected error occurred:", error);
    event.completed({ allowEvent: true });
  }
}

// 注册函数
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);