export const MUTATION_TYPES = Object.freeze({
  INSERT_SLIDE: 'slide.insert',
  REMOVE_SLIDE: 'slide.remove',
  ADD_ELEMENT: 'element.add',
  UPDATE_SHAPE: 'element.updateShape',
  UPDATE_TEXT: 'element.updateText',
  REMOVE_ELEMENT: 'element.remove',
} as const);

export type MutationType = (typeof MUTATION_TYPES)[keyof typeof MUTATION_TYPES];
