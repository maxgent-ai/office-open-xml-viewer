import type { Presentation, ShapeElement, SlideElement } from '@maxgent/ooxml/pptx';

import {
  deriveSlideTreeIndex,
  getElementSources,
  getSlideMutationId,
  insertSlideElement,
  isSlideRegionInsertIndex,
  resolveElementRef,
} from '../adapters/pptx-json-adapter';
import { ELEMENT_ORIGINS } from '../domain/element-origin';
import {
  Mutation,
  type ElementRef,
  type MutationCommandContext,
} from '../domain/mutation';
import { MUTATION_TYPES } from '../domain/mutation-types';
import { MutationExecutionError } from '../engine/errors';
import type { OfficeCliTranslatorError } from '../transport/officecli/errors';
import type { MutationExecutionResult } from '../engine/types';
import {
  OFFICECLI_COMMAND_TYPES,
  OFFICECLI_ELEMENT_TYPES,
} from '../transport/officecli/constants';
import type { OfficeCliCommand } from '../transport/officecli/types';
import { RemoveElementMutation } from './remove-element-mutation';
import {
  freezeProps,
  freezeTarget,
  officeCliError,
  plainTextOf,
  resolveStableSlidePath,
} from './mutation-utils';

export interface AddElementMutationParams {
  readonly target: ElementRef;
  readonly element: SlideElement;
  readonly presentationElementIndex: number;
}

export class AddElementMutation extends Mutation {
  readonly type = MUTATION_TYPES.ADD_ELEMENT;
  readonly target: ElementRef;
  readonly element: SlideElement;
  readonly presentationElementIndex: number;

  constructor({
    target,
    element,
    presentationElementIndex,
  }: AddElementMutationParams) {
    super();
    this.target = freezeTarget(target);
    this.element = dropUnrestorableGeometry(element);
    this.presentationElementIndex = presentationElementIndex;
    Object.freeze(this);
  }

  apply(presentation: Presentation): MutationExecutionResult {
    if (this.target.origin !== ELEMENT_ORIGINS.SLIDE) {
      throw new MutationExecutionError(
        'element.unsupportedOrigin',
        this,
        `Editing ${this.target.origin} elements is not supported`,
      );
    }
    if (resolveElementRef(presentation, this.target)) {
      throw new MutationExecutionError(
        'element.alreadyExists',
        this,
        `Element ${this.target.elementId} already exists`,
      );
    }
    const slideIndex = presentation.slides.findIndex(
      (slide) => getSlideMutationId(slide) === this.target.slideId,
    );
    if (slideIndex < 0) {
      throw new MutationExecutionError(
        'slide.notFound',
        this,
        `Cannot resolve slide ${this.target.slideId}`,
      );
    }
    const slide = presentation.slides[slideIndex];
    const elementSources = getElementSources(slide);
    if (!elementSources) {
      throw new MutationExecutionError(
        'element.metadataUnavailable',
        this,
        `Slide ${this.target.slideId} has no complete element source metadata`,
      );
    }
    if (!isSlideRegionInsertIndex(elementSources, this.presentationElementIndex)) {
      throw new MutationExecutionError(
        'element.invalidIndex',
        this,
        `Cannot insert element at presentation index ${this.presentationElementIndex}`,
      );
    }

    return {
      presentation: insertSlideElement(
        presentation,
        slideIndex,
        this.element,
        this.presentationElementIndex,
      ),
      changedSlideIds: [this.target.slideId],
      changedElements: [this.target],
    };
  }

  inverse(): RemoveElementMutation {
    return new RemoveElementMutation({ target: this.target });
  }

  toOfficeCli(
    presentation: Presentation,
    context: MutationCommandContext,
  ): OfficeCliCommand {
    if (this.element.type !== 'shape') {
      throw officeCliError(
        'target.unsupportedElement',
        context,
        this,
        `OfficeCLI MVP cannot restore ${this.element.type} elements`,
      );
    }
    if (!/^\d+$/.test(this.target.elementId)) {
      throw officeCliError(
        'target.unstableElementId',
        context,
        this,
        `Element ${this.target.elementId} has no stable numeric OOXML id`,
      );
    }
    if (this.element.id && this.element.id !== this.target.elementId) {
      throw officeCliError(
        'target.unstableElementId',
        context,
        this,
        `Element snapshot id ${this.element.id} does not match target id ${this.target.elementId}`,
      );
    }
    const text = this.element.textBody ? plainTextOf(this.element.textBody) : undefined;
    if (this.element.textBody && text === undefined) {
      throw officeCliError(
        'value.invalidText',
        context,
        this,
        'OfficeCLI MVP cannot restore a shape containing math text runs',
      );
    }
    const slideIndex = presentation.slides.findIndex(
      (slide) => getSlideMutationId(slide) === this.target.slideId,
    );
    if (slideIndex < 0) {
      throw officeCliError(
        'target.slideNotFound',
        context,
        this,
        `Cannot resolve slide ${this.target.slideId}`,
      );
    }
    const elementSources = getElementSources(presentation.slides[slideIndex]);
    if (!elementSources) {
      throw officeCliError(
        'target.metadataUnavailable',
        context,
        this,
        `Slide ${this.target.slideId} has no complete element source metadata`,
      );
    }
    if (!isSlideRegionInsertIndex(elementSources, this.presentationElementIndex)) {
      throw officeCliError(
        'value.invalidIndex',
        context,
        this,
        `Cannot insert element at presentation index ${this.presentationElementIndex}`,
      );
    }
    const slideTreeIndex = deriveSlideTreeIndex(
      elementSources,
      this.presentationElementIndex,
    );
    // custGeom and adj values are dropped by dropUnrestorableGeometry in the
    // constructor (degraded restore by design), so no guard is needed here.
    // Per-run text formatting fidelity is not guarded either; plain-text
    // restore keeps the paragraph structure but drops run-level styling.
    const props: Record<string, string> = {
      id: this.target.elementId,
      // OfficeCLI batch 没有顶层插入下标；z-order 是 1-based 的 spTree 位置
      //（zorder=1 在形状树最底层）。由 presentationElementIndex 之前的
      // origin:'slide' 序位推导而来。
      zorder: String(slideTreeIndex + 1),
      preset: this.element.geometry,
      x: `${this.element.x}emu`,
      y: `${this.element.y}emu`,
      width: `${this.element.width}emu`,
      height: `${this.element.height}emu`,
      rotation: String(this.element.rotation),
      flipH: String(this.element.flipH),
      flipV: String(this.element.flipV),
    };
    if (this.element.name) props.name = this.element.name;
    if (text !== undefined) props.text = text;
    this.#applyFillProps(props, context);
    this.#applyStrokeProps(props, context);
    this.#applyShadowProps(props, context);

    return Object.freeze({
      command: OFFICECLI_COMMAND_TYPES.ADD,
      parent: resolveStableSlidePath(presentation, this, context),
      type: OFFICECLI_ELEMENT_TYPES.SHAPE,
      props: freezeProps(props),
    });
  }

  #applyFillProps(props: Record<string, string>, context: MutationCommandContext): void {
    const fill = (this.element as ShapeElement).fill;
    if (fill == null) return;
    switch (fill.fillType) {
      case 'none':
        props.fill = 'none';
        return;
      case 'solid': {
        const color = splitColorAlpha(fill.color);
        if (!color) {
          throw this.#fidelityError(context, `Fill color ${fill.color} is not a plain hex color`);
        }
        props.fill = color.hex;
        if (color.alpha !== undefined) props.opacity = formatFraction(color.alpha);
        return;
      }
      case 'pattern': {
        if (!HEX6_PATTERN.test(fill.fg) || !HEX6_PATTERN.test(fill.bg)) {
          throw this.#fidelityError(
            context,
            'Pattern fill colors with alpha cannot be expressed through the officecli pattern grammar',
          );
        }
        props.pattern = `${fill.preset}:${fill.fg}:${fill.bg}`;
        return;
      }
      case 'gradient': {
        // The officecli gradient grammar covers a two-stop linear ramp with
        // an integer angle; other authored gradients (radial/path focus,
        // extra stops, alpha stops) would round-trip lossily.
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
          throw this.#fidelityError(
            context,
            'Only two-stop linear gradients with an integer angle can be restored through officecli',
          );
        }
        props.gradient = `LINEAR;${start.color};${end.color};${fill.angle}`;
        return;
      }
      default:
        throw this.#fidelityError(
          context,
          'Image fills cannot be restored: the translator has no access to the embedded media bytes',
        );
    }
  }

  #applyStrokeProps(props: Record<string, string>, context: MutationCommandContext): void {
    const stroke = (this.element as ShapeElement).stroke;
    if (stroke == null) return;
    if (stroke.fill) {
      throw this.#fidelityError(
        context,
        'Gradient/pattern outline paint cannot be restored through the officecli line grammar yet',
      );
    }
    const color = splitColorAlpha(stroke.color);
    if (!color) {
      throw this.#fidelityError(context, `Outline color ${stroke.color} is not a plain hex color`);
    }
    props.line = `${color.hex}:${emuToPoints(stroke.width)}`;
    if (color.alpha !== undefined) props.lineOpacity = formatFraction(color.alpha);
    if (stroke.dashStyle) props.lineDash = stroke.dashStyle;
    if (stroke.lineCap) {
      // The model normalizes DrawingML @cap="flat" to the Canvas token "butt".
      props.lineCap = stroke.lineCap === 'butt' ? 'flat' : stroke.lineCap;
    }
    if (stroke.cmpd && stroke.cmpd !== 'sng') props.cmpd = stroke.cmpd;
    this.#applyArrowEnd(props, 'headEnd', stroke.headEnd, context);
    this.#applyArrowEnd(props, 'tailEnd', stroke.tailEnd, context);
  }

  #applyArrowEnd(
    props: Record<string, string>,
    key: 'headEnd' | 'tailEnd',
    arrow: { type: string; w: string; len: string } | undefined,
    context: MutationCommandContext,
  ): void {
    if (!arrow || arrow.type === 'none') return;
    // officecli only takes the arrowhead type; "med" is the OOXML default
    // for both size multipliers, so anything else would restore lossily.
    if (arrow.w !== 'med' || arrow.len !== 'med') {
      throw this.#fidelityError(
        context,
        `Arrowhead ${key} uses non-default size multipliers that officecli cannot express`,
      );
    }
    props[key] = arrow.type;
  }

  #applyShadowProps(props: Record<string, string>, context: MutationCommandContext): void {
    const shadow = (this.element as ShapeElement).shadow;
    if (shadow == null) return;
    if (!HEX6_PATTERN.test(shadow.color)) {
      throw this.#fidelityError(context, `Shadow color ${shadow.color} is not a plain hex color`);
    }
    // officecli compound form: color-blurPt-angleDeg-distPt-opacityPct.
    props.shadow = [
      shadow.color,
      emuToPoints(shadow.blur),
      formatFraction(shadow.dir),
      emuToPoints(shadow.dist),
      formatFraction(shadow.alpha * 100),
    ].join('-');
  }

  #fidelityError(context: MutationCommandContext, message: string): OfficeCliTranslatorError {
    return officeCliError('value.unsupportedFidelity', context, this, message);
  }
}

/**
 * Degrades geometry details that cannot round-trip through OfficeCLI, so the
 * optimistic apply() result stays identical to the authoritative file:
 *
 * - `custGeom` path data has no structured `add shape` property in OfficeCLI;
 *   the shape falls back to a plain preset rectangle.
 * - `adj`..`adj8` adjust values are stored positionally by the parser without
 *   their `<a:gd>` names, while the officecli `adj=name:formula` grammar
 *   needs those names; restoring by position could bind a value to the wrong
 *   handle, so the values are dropped and preset defaults apply. Extending
 *   the parser model to retain gd names would lift this (tracked follow-up).
 */
function dropUnrestorableGeometry(element: SlideElement): SlideElement {
  if (element.type !== 'shape') return element;
  const shape = element as ShapeElement;
  const hasAdjustments = shape.adj != null
    || shape.adj2 != null
    || shape.adj3 != null
    || shape.adj4 != null
    || shape.adj5 != null
    || shape.adj6 != null
    || shape.adj7 != null
    || shape.adj8 != null;
  if (shape.custGeom == null && !hasAdjustments) return element;
  return {
    ...shape,
    geometry: shape.custGeom != null ? 'rect' : shape.geometry,
    custGeom: null,
    adj: null,
    adj2: null,
    adj3: null,
    adj4: null,
    adj5: null,
    adj6: null,
    adj7: null,
    adj8: null,
  };
}

const HEX6_PATTERN = /^[0-9A-Fa-f]{6}$/;
const HEX8_PATTERN = /^[0-9A-Fa-f]{8}$/;

/** Splits an RRGGBB / RRGGBBAA model color into hex + fractional alpha. */
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

/** Formats fractional prop values without float-noise tails (0.50196..., 1.5). */
function formatFraction(value: number): string {
  return String(Number(value.toFixed(6)));
}
