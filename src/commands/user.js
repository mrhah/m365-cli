import graphClient from '../graph/client.js';
import { outputUserSearchResults } from '../utils/output.js';
import { handleError } from '../utils/error.js';

/**
 * Search users and personal contacts
 */
export async function searchUsers(query, options) {
  try {
    const { top = 10, json = false } = options;

    if (!query) {
      throw new Error('Search query is required');
    }

    const results = await graphClient.user.search(query, { top });
    outputUserSearchResults(results, { json, query, top });
  } catch (error) {
    handleError(error, { json: options.json });
  }
}

export default {
  search: searchUsers,
};
