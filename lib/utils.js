const path = require('path');

const sanitizeHtml = (text) => {
  if (typeof text !== 'string') return text;
  return text.replace(/<[^>]*>?/gm, '')
           .replace(/(&lt;|&gt;|&quot;|&#x27;|&amp;)/gi, (match) => {
             const escapeMap = {
               '&lt;': '<',
               '&gt;': '>',
               '&quot;': '"',
               '&#x27;': "'",
               '&amp;': '&'
             };
             return escapeMap[match.toLowerCase()] || match;
           });
};

const isValidEmail = (email) => {
  const emailRegex = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
  return typeof email === 'string' && emailRegex.test(email);
};

const isValidUsername = (username) => {
  if (typeof username !== 'string' || username.length < 3 || username.length > 30) return false;
  const usernameRegex = /^[a-zA-Z0-9 _-]+$/;
  return usernameRegex.test(username);
};

const isValidPassword = (password) => {
  if (typeof password !== 'string' || password.length < 6) return false;
  return true;
};

const isValidId = (id) => {
  const uuidRegex = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
  const numericRegex = /^[0-9]+$/;
  return typeof id === 'string' && (uuidRegex.test(id) || numericRegex.test(id));
};

const sanitizeFilePath = (filePath) => {
  if (typeof filePath !== 'string') return null;
  const resolvedPath = path.resolve(filePath);
  const basePath = path.resolve('.');
  if (!resolvedPath.startsWith(basePath)) {
    return null;
  }
  return resolvedPath;
};

// XSS / injection patterns only — allow common product-name symbols (. / ' " & + % () etc.)
const DANGEROUS_INPUT_PATTERN = /(<\s*\/?\s*script\b|javascript\s*:|vbscript\s*:|data\s*:\s*text\/html|<\s*iframe\b|on(?:load|error|click|mouse\w+|focus|blur)\s*=|eval\s*\(|expression\s*\()/i;

const validateAndSanitizeInput = (input, type = 'general') => {
  if (input === null || input === undefined) return null;

  let sanitizedInput = input;

  if (typeof sanitizedInput === 'string') {
    // Reject XSS payloads before stripping tags so "<script>..." still fails validation
    if (type === 'general' && DANGEROUS_INPUT_PATTERN.test(sanitizedInput)) {
      throw new Error('Input contains potentially dangerous characters');
    }
    sanitizedInput = sanitizeHtml(sanitizedInput);
  }

  switch (type) {
    case 'email':
      if (!isValidEmail(sanitizedInput)) {
        throw new Error('Invalid email format');
      }
      break;
    case 'username':
      if (!isValidUsername(sanitizedInput)) {
        throw new Error('Invalid username format');
      }
      break;
    case 'password':
      if (!isValidPassword(sanitizedInput)) {
        throw new Error('Invalid password format');
      }
      break;
    case 'id':
      if (!isValidId(sanitizedInput)) {
        throw new Error('Invalid ID format');
      }
      break;
    case 'filename':
      if (typeof sanitizedInput !== 'string' || !/^[a-zA-Z0-9._-]+$/.test(sanitizedInput)) {
        throw new Error('Invalid filename format');
      }
      break;
    default:
      // general: HTML already stripped; symbols in names/notes/SKU are allowed
      break;
  }

  return sanitizedInput;
};

module.exports = {
  sanitizeHtml,
  isValidEmail,
  isValidUsername,
  isValidPassword,
  isValidId,
  sanitizeFilePath,
  validateAndSanitizeInput,
};
