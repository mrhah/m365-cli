import { readFileSync, writeFileSync } from 'fs';
import { basename } from 'path';
import graphClient from '../graph/client.js';
import { outputMailList, outputMailDetail, outputSendResult, outputAttachmentList, outputAttachmentDownload } from '../utils/output.js';
import { handleError } from '../utils/error.js';
import { isTrustedSender, addTrustedSender, removeTrustedSender, listTrustedSenders, getWhitelistFilePath } from '../utils/trusted-senders.js';

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
    
    // Mark untrusted senders
    const mailsWithTrustStatus = mails.map(mail => {
      const senderEmail = mail.from?.emailAddress?.address;
      return {
        ...mail,
        isTrusted: senderEmail ? isTrustedSender(senderEmail) : false,
      };
    });
    
    outputMailList(mailsWithTrustStatus, { json, top });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

/**
 * Read email by ID
 */
export async function readMail(id, options) {
  try {
    const { json = false, force = false } = options;
    
    if (!id) {
      throw new Error('Email ID is required');
    }
    
    const mail = await graphClient.mail.get(id);
    
    // Check whitelist unless --force is used
    const senderEmail = mail.from?.emailAddress?.address;
    const trusted = senderEmail ? isTrustedSender(senderEmail) : false;
    
    if (!trusted && !force) {
      // Filter content for untrusted senders
      mail.bodyFiltered = true;
      mail.originalBody = mail.body;
      mail.body = {
        contentType: mail.body?.contentType || 'Text',
        content: '[内容已过滤 - 发件人不在白名单中]\n\n使用 --force 选项可跳过白名单检查。',
      };
    }
    
    mail.isTrusted = trusted;
    
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

/**
 * Manage trusted senders whitelist
 */
export async function trustSender(email, options) {
  try {
    const { json = false } = options;
    
    if (!email) {
      throw new Error('Email address or domain is required');
    }
    
    addTrustedSender(email);
    
    const result = {
      action: 'added',
      entry: email,
      path: getWhitelistFilePath(),
    };
    
    if (json) {
      console.log(JSON.stringify(result, null, 2));
    } else {
      console.log(`✅ Added to whitelist: ${email}`);
      console.log(`   File: ${result.path}`);
    }
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

export async function untrustSender(email, options) {
  try {
    const { json = false } = options;
    
    if (!email) {
      throw new Error('Email address or domain is required');
    }
    
    removeTrustedSender(email);
    
    const result = {
      action: 'removed',
      entry: email,
      path: getWhitelistFilePath(),
    };
    
    if (json) {
      console.log(JSON.stringify(result, null, 2));
    } else {
      console.log(`❌ Removed from whitelist: ${email}`);
      console.log(`   File: ${result.path}`);
    }
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

export async function showTrustedSenders(options) {
  try {
    const { json = false } = options;
    
    const trustedSenders = listTrustedSenders();
    
    if (json) {
      console.log(JSON.stringify({ trustedSenders, path: getWhitelistFilePath() }, null, 2));
    } else {
      console.log(`📋 Trusted senders whitelist:`);
      console.log(`   File: ${getWhitelistFilePath()}`);
      console.log('');
      
      if (trustedSenders.length === 0) {
        console.log('   (empty)');
      } else {
        trustedSenders.forEach(entry => {
          if (entry.startsWith('@')) {
            console.log(`   🌐 ${entry} (domain)`);
          } else {
            console.log(`   📧 ${entry}`);
          }
        });
      }
      
      console.log('');
      console.log(`   Total: ${trustedSenders.length} entries`);
    }
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
  trust: trustSender,
  untrust: untrustSender,
  trusted: showTrustedSenders,
};
