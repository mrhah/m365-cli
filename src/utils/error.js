/**
 * Custom error classes for M365 CLI
 */

export class M365Error extends Error {
  constructor(message, code, details = null) {
    super(message);
    this.name = 'M365Error';
    this.code = code;
    this.details = details;
  }
}

export class AuthError extends M365Error {
  constructor(message, details = null) {
    super(message, 'AUTH_ERROR', details);
    this.name = 'AuthError';
  }
}

export class ApiError extends M365Error {
  constructor(message, statusCode, details = null) {
    super(message, 'API_ERROR', details);
    this.name = 'ApiError';
    this.statusCode = statusCode;
  }
}

export class TokenExpiredError extends AuthError {
  constructor() {
    super('Token expired and refresh failed. Please run: m365 login');
    this.name = 'TokenExpiredError';
  }
}

export class InsufficientPrivilegesError extends ApiError {
  constructor(message, details = null) {
    super(message || 'Insufficient privileges. Additional permissions are required.', 403, details);
    this.name = 'InsufficientPrivilegesError';
    this.code = 'INSUFFICIENT_PRIVILEGES';
  }
}

/**
 * Handle and format errors
 */
export function handleError(error, options = {}) {
  const { json = false } = options;
  
  if (json) {
    const errorObj = {
      error: true,
      message: error.message,
      code: error.code || 'UNKNOWN_ERROR',
    };
    
    if (error.details) {
      errorObj.details = error.details;
    }
    
    if (error.statusCode) {
      errorObj.statusCode = error.statusCode;
    }
    
    console.error(JSON.stringify(errorObj, null, 2));
  } else {
    console.error(`❌ Error: ${error.message}`);
    
    if (error.details) {
      console.error(`   Details: ${JSON.stringify(error.details, null, 2)}`);
    }
  }
  
  process.exit(1);
}

/**
 * Parse Graph API error response
 */
export function parseGraphError(response, statusCode) {
  let message = `API Error (${statusCode})`;
  let details = null;
  
  if (response.error) {
    const errorCode = response.error.code;
    
    // Map common error codes to user-friendly messages
    const friendlyMessages = {
      'ErrorInvalidIdMalformed': '无效的 ID 格式。请检查您提供的 ID 是否正确。',
      'itemNotFound': '找不到指定的文件或文件夹。请检查路径是否存在。',
      'InvalidRequest': '无效的请求。请检查您的参数格式。',
      'invalidRequest': '无效的请求。请检查您的参数格式。',
      'InvalidHostname': '无效的站点主机名。请检查 SharePoint 站点格式。',
      'resourceNotFound': '找不到指定的资源。',
      'ErrorItemNotFound': '找不到指定的项目。',
      'mailboxNotFound': '找不到邮箱。',
      'folderNotFound': '找不到指定的文件夹。',
      'ErrorNonExistentMailbox': '邮箱不存在。',
    };
    
    // Use friendly message if available, otherwise use original
    message = friendlyMessages[errorCode] || response.error.message || message;
    
    // Include error code in details for debugging
    details = {
      code: errorCode,
      originalMessage: response.error.message,
      innerError: response.error.innerError,
    };
  }
  
  // Detect insufficient privileges (scope issue) — return specialized error
  if (statusCode === 403 && response.error) {
    const errorCode = response.error.code;
    const innerCode = response.error.innerError?.code;
    const insufficientCodes = [
      'Authorization_RequestDenied',
      'AccessDenied',
      'ErrorAccessDenied',
      'InsufficientPrivileges',
    ];
    
    if (insufficientCodes.includes(errorCode) || insufficientCodes.includes(innerCode)) {
      return new InsufficientPrivilegesError(message, details);
    }
  }
  
  return new ApiError(message, statusCode, details);
}

export default {
  M365Error,
  AuthError,
  ApiError,
  TokenExpiredError,
  InsufficientPrivilegesError,
  handleError,
  parseGraphError,
};
