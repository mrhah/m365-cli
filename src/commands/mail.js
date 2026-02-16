import { readFileSync, writeFileSync } from 'fs';
import { basename } from 'path';
import graphClient from '../graph/client.js';
import { outputMailList, outputMailDetail, outputSendResult, outputAttachmentList, outputAttachmentDownload } from '../utils/output.js';
import { handleError } from '../utils/error.js';

/**
 * Mail commands
 */

/**
 * List emails
 */
export async function listMails(options) {
  try {
    const { top = 10, folder = 'inbox', json = false } = options;
    
    const mails = await graphClient.mail.list({ top, folder });
    
    outputMailList(mails, { json, top });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Read email by ID
 */
export async function readMail(id, options) {
  try {
    const { json = false } = options;
    
    if (!id) {
      throw new Error('Email ID is required');
    }
    
    const mail = await graphClient.mail.get(id);
    
    outputMailDetail(mail, { json });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Send email
 */
export async function sendMail(to, subject, body, options) {
  try {
    const { attach = [], json = false } = options;
    
    if (!to || !subject || !body) {
      throw new Error('To, subject, and body are required');
    }
    
    // Build message
    const message = {
      subject,
      body: {
        contentType: 'HTML',
        content: body,
      },
      toRecipients: [
        {
          emailAddress: {
            address: to,
          },
        },
      ],
    };
    
    // Handle attachments
    if (attach && attach.length > 0) {
      message.attachments = [];
      
      let totalSize = 0;
      
      for (const filePath of attach) {
        try {
          const fileBuffer = readFileSync(filePath);
          const fileName = filePath.split('/').pop();
          const base64Content = fileBuffer.toString('base64');
          
          totalSize += fileBuffer.length;
          
          message.attachments.push({
            '@odata.type': '#microsoft.graph.fileAttachment',
            name: fileName,
            contentBytes: base64Content,
          });
        } catch (error) {
          throw new Error(`Failed to read attachment: ${filePath} - ${error.message}`);
        }
      }
      
      // Warn about size limit
      if (totalSize > 2360320) { // ~2.25MB
        console.warn('⚠️  Warning: Total attachment size exceeds recommended limit (~2.25MB).');
        console.warn('   Large emails may fail due to Graph API limits.');
      }
    }
    
    // Send email
    await graphClient.mail.send(message);
    
    const result = {
      status: 'sent',
      to,
      subject,
      attachments: attach.length,
    };
    
    outputSendResult(result, { json });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Search emails
 */
export async function searchMails(query, options) {
  try {
    const { top = 10, json = false } = options;
    
    if (!query) {
      throw new Error('Search query is required');
    }
    
    const mails = await graphClient.mail.search(query, { top });
    
    if (!json) {
      console.log(`🔍 Search results for: "${query}"`);
      console.log('');
    }
    
    outputMailList(mails, { json, top });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * List attachments for an email
 */
export async function listAttachments(id, options) {
  try {
    const { json = false } = options;
    
    if (!id) {
      throw new Error('Email ID is required');
    }
    
    const attachments = await graphClient.mail.attachments(id);
    
    outputAttachmentList(attachments, { json, messageId: id });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Download email attachment
 */
export async function downloadAttachment(messageId, attachmentId, localPath, options) {
  try {
    const { json = false } = options;
    
    if (!messageId || !attachmentId) {
      throw new Error('Message ID and Attachment ID are required');
    }
    
    // Get attachment data
    const attachment = await graphClient.mail.downloadAttachment(messageId, attachmentId);
    
    if (!attachment.contentBytes) {
      throw new Error('Attachment content not found');
    }
    
    // Determine output path
    const fileName = attachment.name || 'attachment';
    const outputPath = localPath || fileName;
    
    // Decode base64 content and save
    const buffer = Buffer.from(attachment.contentBytes, 'base64');
    writeFileSync(outputPath, buffer);
    
    const result = {
      name: fileName,
      path: outputPath,
      size: buffer.length,
    };
    
    outputAttachmentDownload(result, { json });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

export default {
  list: listMails,
  read: readMail,
  send: sendMail,
  search: searchMails,
  attachments: listAttachments,
  downloadAttachment,
};
