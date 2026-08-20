import type { Fill, Stroke } from '@maxgent/ooxml/pptx';

import type { Mutation, MutationCommandContext } from '../domain/mutation';
import type { OfficeCliTranslatorError } from '../transport/officecli/errors';
import { officeCliError } from './mutation-utils';

const HEX6_PATTERN = /^[0-9A-Fa-f]{6}$/;
const HEX8_PATTERN = /^[0-9A-Fa-f]{8}$/;

export function applyShapeFillProps(
  props: Record<string, string>,
  fill: Fill,
  context: MutationCommandContext,
  mutation: Mutation,
): void {
  switch (fill.fillType) {
    case 'none':
      props.fill = 'none';
      return;
    case 'solid': {
      const color = splitColorAlpha(fill.color);
      if (!color) {
        throw fidelityError(context, mutation, `Fill color ${fill.color} is not a plain hex color`);
      }
      props.fill = color.hex;
      if (color.alpha !== undefined) props.opacity = formatFraction(color.alpha);
      return;
    }
    case 'pattern':
      if (!HEX6_PATTERN.test(fill.fg) || !HEX6_PATTERN.test(fill.bg)) {
        throw fidelityError(
          context,
          mutation,
          'Pattern fill colors with alpha cannot be expressed through the officecli pattern grammar',
        );
      }
      props.pattern = `${fill.preset}:${fill.fg}:${fill.bg}`;
      return;
    case 'gradient': {
      // OfficeCLI can round-trip only this strict subset of DrawingML gradients.
      const [start, end] = fill.stops;
      if (
        fill.gradType !== 'linear'
        || fill.stops.length !== 2
        || start.position !== 0
        || end.position !== 1
        || !HEX6_PATTERN.test(start.color)
        || !HEX6_PATTERN.test(end.color)
        || !Number.isInteger(fill.angle)
      ) {
        throw fidelityError(
          context,
          mutation,
          'Only two-stop linear gradients with an integer angle can be restored through officecli',
        );
      }
      props.gradient = `LINEAR;${start.color};${end.color};${fill.angle}`;
      return;
    }
    default:
      throw fidelityError(
        context,
        mutation,
        'Image fills cannot be restored: the translator has no access to the embedded media bytes',
      );
  }
}

export function applyShapeStrokeProps(
  props: Record<string, string>,
  stroke: Stroke,
  context: MutationCommandContext,
  mutation: Mutation,
): void {
  if (stroke.fill) {
    throw fidelityError(
      context,
      mutation,
      'Gradient/pattern outline paint cannot be restored through the officecli line grammar yet',
    );
  }
  const color = splitColorAlpha(stroke.color);
  if (!color) {
    throw fidelityError(context, mutation, `Outline color ${stroke.color} is not a plain hex color`);
  }
  props.line = `${color.hex}:${emuToPoints(stroke.width)}`;
  if (color.alpha !== undefined) props.lineOpacity = formatFraction(color.alpha);
  if (stroke.dashStyle) props.lineDash = stroke.dashStyle;
  if (stroke.lineCap) props.lineCap = stroke.lineCap === 'butt' ? 'flat' : stroke.lineCap;
  if (stroke.cmpd && stroke.cmpd !== 'sng') props.cmpd = stroke.cmpd;
  applyArrowEnd(props, 'headEnd', stroke.headEnd, context, mutation);
  applyArrowEnd(props, 'tailEnd', stroke.tailEnd, context, mutation);
}

function applyArrowEnd(
  props: Record<string, string>,
  key: 'headEnd' | 'tailEnd',
  arrow: { type: string; w: string; len: string } | undefined,
  context: MutationCommandContext,
  mutation: Mutation,
): void {
  if (!arrow || arrow.type === 'none') return;
  if (arrow.w !== 'med' || arrow.len !== 'med') {
    throw fidelityError(
      context,
      mutation,
      `Arrowhead ${key} uses non-default size multipliers that officecli cannot express`,
    );
  }
  props[key] = arrow.type;
}

function fidelityError(
  context: MutationCommandContext,
  mutation: Mutation,
  message: string,
): OfficeCliTranslatorError {
  return officeCliError('value.unsupportedFidelity', context, mutation, message);
}

function splitColorAlpha(color: string): { hex: string; alpha?: number } | undefined {
  if (HEX6_PATTERN.test(color)) return { hex: color };
  if (HEX8_PATTERN.test(color)) {
    return { hex: color.slice(0, 6), alpha: Number.parseInt(color.slice(6), 16) / 255 };
  }
  return undefined;
}

function emuToPoints(value: number): string {
  return formatFraction(value / 12_700);
}

export function formatFraction(value: number): string {
  return String(Number(value.toFixed(6)));
}
