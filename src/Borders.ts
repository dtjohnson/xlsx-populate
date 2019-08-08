import _ from 'lodash';
import { Border, BorderStyle, DiagonalBorder, DiagonalBorderDirection } from './types';
import { getColor, setColor } from './utils';
import { INode } from './XmlParser';
import * as xmlq from './xmlq';

type BorderSide = 'top'|'bottom'|'left'|'right'|'diagonal';

/**
 * The borders of a cell.
 */
export class Borders {
    public constructor(
        private readonly node: INode,
    ) {}

    public get top(): Border|undefined {
        return this.getBorder('top');
    }   

    public set top(border: Border|undefined) {
        this.setBorder('top', border);
    }

    public get bottom(): Border|undefined {
        return this.getBorder('bottom');
    }   

    public set bottom(border: Border|undefined) {
        this.setBorder('bottom', border);
    }

    public get left(): Border|undefined {
        return this.getBorder('left');
    }   

    public set left(border: Border|undefined) {
        this.setBorder('left', border);
    }

    public get right(): Border|undefined {
        return this.getBorder('right');
    }   

    public set right(border: Border|undefined) {
        this.setBorder('right', border);
    }

    public get diagonal(): DiagonalBorder|undefined {
        const border = this.getBorder('diagonal') as DiagonalBorder;
        const up = this.node.attributes && this.node.attributes.diagonalUp;
        const down = this.node.attributes && this.node.attributes.diagonalDown;
        let direction: DiagonalBorderDirection|undefined;
        if (up && down) direction = 'both';
        else if (up) direction = 'up';
        else if (down) direction = 'down';
        if (direction) border.direction = direction;
        return border;
    }   

    public set diagonal(border: DiagonalBorder|undefined) {
        this.setBorder('diagonal', border);

        if (border) {
            xmlq.setAttributes(this.node, {
                diagonalUp: border.direction === 'up' || border.direction === 'both' ? 1 : undefined,
                diagonalDown: border.direction === 'down' || border.direction === 'both' ? 1 : undefined,
            });
        } else {
            xmlq.setAttributes(this.node, {
                diagonalUp: undefined,
                diagonalDown: undefined,
            });
        }
    }

    private getBorder(side: BorderSide): Border|undefined {
        const node = xmlq.findChild(this.node, side);
        if (!node) return;

        const border: Border = {};

        const style = xmlq.getChildAttribute(this.node, side, 'style') as BorderStyle;
        if (style) border.style = style;
        const color = getColor(node, 'color');
        if (color) border.color = color;
        
        if (_.isEmpty(border)) return;

        return border;
    }

    private setBorder(side: BorderSide, border: Border|undefined): void {
        if (_.isNil(border)) {
            border = { style: undefined, color: undefined };
        }

        if (border && 'style' in border) {
            xmlq.setChildAttributes(this.node, side, { style: border.style });
        }

        if (border && 'color' in border) {
            const node = xmlq.findChild(this.node, side);
            setColor(node!, 'color', border.color);
        }
    }
}
