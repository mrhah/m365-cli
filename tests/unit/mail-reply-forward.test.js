import { describe, it, expect, vi, beforeEach } from 'vitest';

const mockReply = vi.fn();
const mockReplyAll = vi.fn();
const mockForward = vi.fn();

vi.mock('../../src/graph/client.js', () => ({
  default: {
    mail: {
      reply: (...args) => mockReply(...args),
      replyAll: (...args) => mockReplyAll(...args),
      forward: (...args) => mockForward(...args),
    },
    getCurrentUser: vi.fn().mockResolvedValue({
      mail: 'test@test.com',
      userPrincipalName: 'test@test.com',
    }),
  },
}));

const mockOutputMailReplyResult = vi.fn();
const mockOutputMailForwardResult = vi.fn();

vi.mock('../../src/utils/output.js', () => ({
  outputMailList: vi.fn(),
  outputMailDetail: vi.fn(),
  outputSendResult: vi.fn(),
  outputAttachmentList: vi.fn(),
  outputAttachmentDownload: vi.fn(),
  outputMailDeleteResult: vi.fn(),
  outputMailMoveResult: vi.fn(),
  outputMailReplyResult: (...args) => mockOutputMailReplyResult(...args),
  outputMailForwardResult: (...args) => mockOutputMailForwardResult(...args),
  outputMailFolderList: vi.fn(),
  outputMailFolderResult: vi.fn(),
}));

const mockHandleError = vi.fn();
vi.mock('../../src/utils/error.js', () => ({
  handleError: (...args) => mockHandleError(...args),
}));

vi.mock('../../src/utils/trusted-senders.js', () => ({
  isTrustedSender: vi.fn().mockReturnValue(true),
  addTrustedSender: vi.fn(),
  removeTrustedSender: vi.fn(),
  listTrustedSenders: vi.fn().mockReturnValue([]),
  getWhitelistFilePath: vi.fn(),
}));

import { replyMail, replyAllMail, forwardMail } from '../../src/commands/mail.js';

describe('Mail reply/reply-all/forward commands', () => {
  beforeEach(() => {
    vi.clearAllMocks();
    mockReply.mockResolvedValue({});
    mockReplyAll.mockResolvedValue({});
    mockForward.mockResolvedValue({});
  });

  describe('replyMail', () => {
    it('should build comment payload by default', async () => {
      await replyMail('msg-1', 'Thanks for the update', { json: false });

      expect(mockReply).toHaveBeenCalledTimes(1);
      expect(mockReply).toHaveBeenCalledWith('msg-1', {
        comment: 'Thanks for the update',
      });
      expect(mockOutputMailReplyResult).toHaveBeenCalledWith(
        expect.objectContaining({
          status: 'sent',
          action: 'reply',
          id: 'msg-1',
        }),
        { json: false }
      );
    });

    it('should build HTML payload when html is true', async () => {
      await replyMail('msg-1', '<b>Thanks</b>', { html: true, json: true });

      expect(mockReply).toHaveBeenCalledWith('msg-1', {
        message: {
          body: {
            contentType: 'HTML',
            content: '<b>Thanks</b>',
          },
        },
      });
      expect(mockOutputMailReplyResult).toHaveBeenCalledWith(
        expect.objectContaining({ action: 'reply' }),
        { json: true }
      );
    });

    it('should validate missing id', async () => {
      await replyMail('', 'content', { json: false });

      expect(mockReply).not.toHaveBeenCalled();
      expect(mockHandleError).toHaveBeenCalledTimes(1);
      expect(mockHandleError).toHaveBeenCalledWith(expect.any(Error), { json: false });
    });

    it('should validate missing content', async () => {
      await replyMail('msg-1', '', { json: false });

      expect(mockReply).not.toHaveBeenCalled();
      expect(mockHandleError).toHaveBeenCalledTimes(1);
      expect(mockHandleError).toHaveBeenCalledWith(expect.any(Error), { json: false });
    });

    it('should route graph errors through handleError', async () => {
      const error = new Error('Reply failed');
      mockReply.mockRejectedValue(error);

      await replyMail('msg-1', 'content', { json: true });

      expect(mockHandleError).toHaveBeenCalledWith(error, { json: true });
    });
  });

  describe('replyAllMail', () => {
    it('should build comment payload correctly', async () => {
      await replyAllMail('msg-2', 'Replying all', { json: false });

      expect(mockReplyAll).toHaveBeenCalledTimes(1);
      expect(mockReplyAll).toHaveBeenCalledWith('msg-2', {
        comment: 'Replying all',
      });
      expect(mockOutputMailReplyResult).toHaveBeenCalledWith(
        expect.objectContaining({
          status: 'sent',
          action: 'reply-all',
          id: 'msg-2',
        }),
        { json: false }
      );
    });

    it('should validate missing id/content', async () => {
      await replyAllMail('', 'content', { json: false });
      await replyAllMail('msg-2', '', { json: false });

      expect(mockReplyAll).not.toHaveBeenCalled();
      expect(mockHandleError).toHaveBeenCalledTimes(2);
    });
  });

  describe('forwardMail', () => {
    it('should build forward payload with toRecipients', async () => {
      await forwardMail('msg-3', 'alice@example.com', 'FYI', { json: false });

      expect(mockForward).toHaveBeenCalledTimes(1);
      expect(mockForward).toHaveBeenCalledWith('msg-3', {
        comment: 'FYI',
        toRecipients: [
          {
            emailAddress: {
              address: 'alice@example.com',
            },
          },
        ],
      });
      expect(mockOutputMailForwardResult).toHaveBeenCalledWith(
        expect.objectContaining({
          status: 'sent',
          action: 'forward',
          id: 'msg-3',
          to: 'alice@example.com',
          recipientCount: 1,
        }),
        { json: false }
      );
    });

    it('should handle comma-separated recipients', async () => {
      await forwardMail('msg-3', 'alice@example.com, bob@example.com', 'FYI', { json: false });

      expect(mockForward).toHaveBeenCalledWith('msg-3', {
        comment: 'FYI',
        toRecipients: [
          {
            emailAddress: {
              address: 'alice@example.com',
            },
          },
          {
            emailAddress: {
              address: 'bob@example.com',
            },
          },
        ],
      });
    });

    it('should send empty comment when no comment is provided', async () => {
      await forwardMail('msg-3', 'alice@example.com', '', { json: false });

      expect(mockForward).toHaveBeenCalledWith('msg-3', {
        comment: '',
        toRecipients: [
          {
            emailAddress: {
              address: 'alice@example.com',
            },
          },
        ],
      });
    });

    it('should build HTML payload when html is true', async () => {
      await forwardMail('msg-3', 'alice@example.com', '<p>FYI</p>', { html: true, json: true });

      expect(mockForward).toHaveBeenCalledWith('msg-3', {
        message: {
          body: {
            contentType: 'HTML',
            content: '<p>FYI</p>',
          },
        },
        toRecipients: [
          {
            emailAddress: {
              address: 'alice@example.com',
            },
          },
        ],
      });
    });

    it('should validate missing id and recipients', async () => {
      await forwardMail('', 'alice@example.com', 'fwd', { json: false });
      await forwardMail('msg-3', '', 'fwd', { json: false });
      await forwardMail('msg-3', ' , ', 'fwd', { json: false });

      expect(mockForward).not.toHaveBeenCalled();
      expect(mockHandleError).toHaveBeenCalledTimes(3);
    });

    it('should route forward errors through handleError', async () => {
      const error = new Error('Forward failed');
      mockForward.mockRejectedValue(error);

      await forwardMail('msg-3', 'alice@example.com', 'fwd', { json: true });

      expect(mockHandleError).toHaveBeenCalledWith(error, { json: true });
    });
  });
});
