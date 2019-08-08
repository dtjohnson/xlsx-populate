export type VerticalAlignment = 'top'|'center'|'bottom'|'justify'|'distributed';

export type HorizontalAlignment = 'left'|'center'|'right'|'fill'|'justify'|'centerContinuous'|'distributed';

export type FontScheme = 'minor'|'major'|'none';

export type TextDirection = 'left-to-right'|'right-to-left';

export type FontGenericFamily = 'serif'|'sans-serif'|'monospace';

export interface RGBColor {
    rgb: string;
    tint?: number;
}

export interface ThemeColor {
    theme: number;
    tint?: number;
}

export type Color = RGBColor|ThemeColor;

export function isRGBColor(c: Color|undefined): c is ThemeColor {
    return c !== undefined && 'rgb' in c;
}

export function isThemeColor(c: Color|undefined): c is ThemeColor {
    return c !== undefined && 'theme' in c;
}

export type FillPattern = 'gray125'|'darkGray'|'mediumGray'|'lightGray'|'gray0625'|'darkHorizontal'|'darkVertical'|'darkDown'|
    'darkUp'|'darkGrid'|'darkTrellis'|'lightHorizontal'|'lightVertical'|'lightDown'|'lightUp'|'lightGrid'|'lightTrellis';

export interface SolidFill {
    type: 'solid';
    color: Color;
}

export interface PatternFill {
    type: 'pattern';
    pattern: FillPattern;
    foreground: Color;
    background: Color;
}

export interface GradientStop {
    position: number;
    color: Color;
}

export type GradientType = 'linear'|'path';

export interface GradientFill {
    type: 'gradient';
    gradientType: GradientType;
    stops: GradientStop[];
    angle?: number;
    left?: number;
    right?: number;
    top?: number;
    bottom?: number;
}

export type Fill = SolidFill|PatternFill|GradientFill;

export type BorderStyle = 'hair'|'dotted'|'dashDotDot'|'dashed'|'mediumDashDotDot'|'thin'|'slantDashDot'|'mediumDashDot'|'mediumDashed'|'medium'|'thick'|'double';

export interface Border {
    style?: BorderStyle;
    color?: Color;
}

export type DiagonalBorderDirection = 'up'|'down'|'both';

export interface DiagonalBorder extends Border {
    direction: DiagonalBorderDirection;
}

export interface NumberFormatSource {
    getNumberFormatCode(id: number): string;

    getNumberFormatId(code: string): number;
}
