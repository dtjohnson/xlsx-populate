import _ from 'lodash';
import { MAX_COLUMNS, MAX_ROWS } from './consts';
import { INode } from './XmlParser';

/**
 * Page Breaks
 */
export class PageBreaks {
    /**
     * The XML node representing the page breaks.
     */
    private readonly node: INode;

    /**
     * Creates a new instance of PageBreaks
     * @param isHorizontal - True if horizontal (row) breaks, false if vertical (column) breaks
     * @param node - The node.
     * @hidden
     */
    public constructor(
        private readonly isHorizontal: boolean,
        node?: INode,
    ) {
        this.node = node || {
            name: this.isHorizontal ? 'rowBreaks' : 'colBreaks',
        };
    }

    /**
     * Add page-breaks by row/column id
     * @param id - The row or column number
     * @return The page breaks
     */
    public add(id: number): PageBreaks {
        if (!this.node.children) this.node.children = [];
        this.node.children.push({
            name: 'brk',
            attributes: {
                id,
                max: this.isHorizontal ? MAX_COLUMNS - 1 : MAX_ROWS - 1,
                man: 1,
            },
        });

        if (!this.node.attributes) this.node.attributes = {};
        this.node.attributes.count = Number(this.node.attributes.count || 0) + 1;
        this.node.attributes.manualBreakCount = Number(this.node.attributes.manualBreakCount || 0) + 1;

        return this;
    }

    /**
     * Remove page-breaks by index
     * @param id - The row or column number
     * @return The page-breaks
     */
    public remove(id: number): PageBreaks {
        if (!this.node.children) return this;
        const brkIndex = _.findIndex(
            this.node.children,
            brk => typeof brk !== 'string' && typeof brk !== 'number' && !!brk.attributes && brk.attributes.id === id,
        );
        
        if (brkIndex >= 0) {
            const brk = this.node.children[brkIndex];
            this.node.children.splice(brkIndex, 1);
            if (!this.node.attributes) this.node.attributes = {};
            this.node.attributes.count = Number(this.node.attributes.count || 0) - 1;
            if (typeof brk !== 'string' && typeof brk !== 'number' && brk.attributes && brk.attributes.man) {
                this.node.attributes.manualBreakCount = Number(this.node.attributes.manualBreakCount || 0) - 1;
            }
        }

        return this;
    }

    /**
     * Get list of page-breaks
     * @return list of the page-breaks
     */
    public list(): number[] {
        const res: number[] = [];
        _.forEach(this.node.children, brk => {
            if (typeof brk !== 'string' && typeof brk !== 'number' && brk.attributes) {
                res.push(Number(brk.attributes.id));
            }
        });

        return res;
    }

    /**
     * Convert the collection to an XML object.
     * @returns The XML or undefined if empty.
     * @hidden
     */
    public toXml(): INode|undefined {
        if (!this.node.children || !this.node.children.length) return;
        return this.node;
    }
}

// tslint:disable
/*
<rowBreaks count="2" manualBreakCount="2">
    <brk id="8" max="16383" man="1"/>
    <brk id="15" max="16383" man="1"/>
</rowBreaks>
<colBreaks count="2" manualBreakCount="2">
    <brk id="4" max="1048575" man="1"/>
    <brk id="8" max="1048575" man="1"/>
</colBreaks>
*/