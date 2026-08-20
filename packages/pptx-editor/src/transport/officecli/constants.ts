export const OFFICECLI_BATCH_SCHEMA_VERSION = '1.0';
export const OFFICECLI_VERSION = '1.0.139';

export const OFFICECLI_COMMAND_TYPES = Object.freeze({
  ADD: 'add',
  SET: 'set',
  REMOVE: 'remove',
} as const);

export const OFFICECLI_ELEMENT_TYPES = Object.freeze({
  SLIDE: 'slide',
  SHAPE: 'shape',
} as const);
