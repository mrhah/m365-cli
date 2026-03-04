import { describe, it, expect, vi, beforeEach } from 'vitest';

const mockUserSearch = vi.fn();
const mockOutputUserSearchResults = vi.fn();
const mockHandleError = vi.fn();

vi.mock('../src/graph/client.js', () => ({
  default: {
    user: {
      search: (...args) => mockUserSearch(...args),
    },
  },
}));

vi.mock('../src/utils/output.js', () => ({
  outputUserSearchResults: (...args) => mockOutputUserSearchResults(...args),
}));

vi.mock('../src/utils/error.js', () => ({
  handleError: (...args) => mockHandleError(...args),
}));

import { searchUsers } from '../src/commands/user.js';

describe('User search command', () => {
  beforeEach(() => {
    vi.clearAllMocks();
  });

  it('searches users and contacts and outputs results', async () => {
    const results = [
      { name: 'Jerry Adams', email: 'jerry@contoso.com', source: 'organization' },
      { name: 'Jerry Personal', email: 'jerry.personal@example.com', source: 'contact' },
    ];
    mockUserSearch.mockResolvedValue(results);

    await searchUsers('Jerry', { top: 5, json: false });

    expect(mockUserSearch).toHaveBeenCalledWith('Jerry', { top: 5 });
    expect(mockOutputUserSearchResults).toHaveBeenCalledWith(results, { json: false, query: 'Jerry', top: 5 });
    expect(mockHandleError).not.toHaveBeenCalled();
  });

  it('uses default top value when not provided', async () => {
    mockUserSearch.mockResolvedValue([]);

    await searchUsers('Alice', { json: true });

    expect(mockUserSearch).toHaveBeenCalledWith('Alice', { top: 10 });
    expect(mockOutputUserSearchResults).toHaveBeenCalledWith([], { json: true, query: 'Alice', top: 10 });
  });

  it('handles missing query as an error', async () => {
    await searchUsers('', { json: true });

    expect(mockUserSearch).not.toHaveBeenCalled();
    expect(mockHandleError).toHaveBeenCalledTimes(1);
    expect(mockHandleError.mock.calls[0][0].message).toBe('Search query is required');
    expect(mockHandleError).toHaveBeenCalledWith(expect.any(Error), { json: true });
  });
});
