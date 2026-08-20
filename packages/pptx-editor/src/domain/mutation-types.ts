export const MUTATION_TYPES = Object.freeze({
  ADD_ELEMENT: 'element.add',
  UPDATE_SHAPE: 'element.updateShape',
  UPDATE_TRANSFORM: 'element.updateTransform',
  UPDATE_TEXT: 'element.updateText',
  REMOVE_ELEMENT: 'element.remove',
} as const);

export type MutationType = (typeof MUTATION_TYPES)[keyof typeof MUTATION_TYPES];
