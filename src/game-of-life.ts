// The board dimension (number of columns/rows).
const BOARD_WIDTH = 240;
const BOARD_HEIGHT = 135;

// How many generations you want to evolve.
const NUMBER_OF_GENERATIONS = 120;

// Refer to https://copy.sh/life/examples for more sample patterns.
const PATTERN_URL = "https://copy.sh/life/examples/glider.rle";

async function main(workbook: ExcelScript.Workbook): Promise<void> {
    let sheet = workbook.addWorksheet();
    sheet.activate();

    const pattern = await Pattern.fromUrl(PATTERN_URL);
    const board = new Board(BOARD_WIDTH, BOARD_HEIGHT, pattern);
    const renderer = new Renderer(sheet);

    console.log("Initializing the world...")
    renderer.initializeCanvas(board.width, board.height);

    console.log("Starting first generation...");
    renderer.renderEvolution(board.getInitialEvolution());

    console.log("Evolving...");

    // Rendering might fail if the interval is too small. Normally it'd be fine if >= 500 milliseconds.
    const RENDER_INTERVAL_MILLISECONDS = 500;

    for (var generation = 1; generation < NUMBER_OF_GENERATIONS; generation++) {
        await sleep(RENDER_INTERVAL_MILLISECONDS);
        let evolution = board.evolveOneGeneration();
        renderer.renderEvolution(evolution);
    }
}

interface Evolution {
    evolvedCells: [number, number, boolean][];
}

interface Grid {
    width: number;
    height: number;
    matrix: boolean[][];
};

class Board implements Grid {
    readonly matrix: boolean[][];

    constructor(public readonly width: number, public readonly height: number, public readonly initialPattern: Pattern) {
        this.matrix = new Array(height).fill(false).map(() => new Array(width).fill(false));
    }

    getInitialEvolution(): Evolution {
        let evolution = {
            evolvedCells: new Array<[number, number, boolean]>()
        }

        let patternX = Math.floor((this.width - this.initialPattern.width) / 2);
        let patternY = Math.floor((this.height - this.initialPattern.height) / 2);
        for (var y = 0; y < this.initialPattern.height; y++) {
            for (var x = 0; x < this.initialPattern.width; x++) {
                this.matrix[y + patternY][x + patternX] = this.initialPattern.matrix[y][x];
                if (this.initialPattern.matrix[y][x]) {
                    evolution.evolvedCells.push([y + patternY, x + patternX, true]);
                }
            }
        }

        return evolution;
    }

    evolveOneGeneration(): Evolution {
        let evolution = {
            evolvedCells: new Array<[number, number, boolean]>()
        };

        for (var y = 0; y < this.height; y++) {
            for (var x = 0; x < this.width; x++) {
                this.evolveCell(y, x, evolution);
            }
        }

        evolution.evolvedCells.forEach(cell => this.matrix[cell[0]][cell[1]] = cell[2]);

        return evolution;
    }

    private evolveCell(y: number, x: number, evolution: Evolution): void {
        const neighbors = this.countCellNeighbors(y, x);
        const previouslyAlive = this.matrix[y][x];
        switch (true) {
            case (previouslyAlive && neighbors < 2): evolution.evolvedCells.push([y, x, false]); break;
            case (previouslyAlive && (neighbors === 2 || neighbors == 3)): break;
            case (previouslyAlive && neighbors > 3): evolution.evolvedCells.push([y, x, false]); break;
            case (!previouslyAlive && neighbors === 3): evolution.evolvedCells.push([y, x, true]); break;
            default: break;
        }
    }

    private countCellNeighbors(cellY: number, cellX: number): number {
        let count = 0;
        for (let x = -1; x <= 1; x++) {
            for (let y = -1; y <= 1; y++) {
                if (x === 0 && y === 0) continue;
                const posX = cellX + x;
                const posY = cellY + y;
                if (posY >= 0 && posY < this.height && posX >= 0 && posX < this.width) {
                    if (this.matrix[posY][posX]) {
                        count++;
                    }
                }
            }
        }

        return count;
    }
}

class Pattern implements Grid {
    width: number;
    height: number;
    matrix: boolean[][];
    name: string;
    rule: string;

    static async fromUrl(url: string): Promise<Pattern> {
        let fetchResult = await fetch(`https://sofetch.glitch.me/${encodeURI(url)}`);
        let patternFileContent = await fetchResult.text();
        let lines = patternFileContent.split("\n");
        let pattern: Pattern = { width: 0, height: 0, matrix: null, name: "", rule: "" };
        let patternString = "";
        lines.forEach(line => {
            if (line.toUpperCase().startsWith("#N ")) {
                pattern.name = /#N (.+)/.exec(line)[1];
            } else if (line.startsWith("#")) {
                // Ignore for now
            } else if (line.startsWith("x")) {
                const regex = /x = (?<width>\d+), y = (?<height>\d+), rule = (?<rule>.+)/;
                const matchGroups = regex.exec(line).groups;
                pattern.width = +matchGroups.width;
                pattern.height = +matchGroups.height;
                pattern.rule = matchGroups.rule;
            } else {
                patternString += line;
            }
        });

        pattern.matrix = Pattern.parsePattern(patternString.replace("!", ""), pattern.width, pattern.height);

        return pattern;
    }

    private constructor() {
    }

    private static parsePattern(patternString: string, width: number, height: number): boolean[][] {
        let matrix: boolean[][] =
            new Array(height).fill(false)
                .map(() => new Array(width).fill(false));

        const rows = patternString.split("$");
        const regex = /(\d*)([bo])/g;
        for (var y = 0; y < height; y++) {
            const row = rows[y];
            let matchElement: RegExpExecArray = null;
            let x = 0;
            do {
                matchElement = regex.exec(row);
                if (!matchElement) {
                    continue;
                }

                const matchLength = matchElement[1] ? matchElement[1] : 1;
                const alive = matchElement[2] === "o";
                for (let index = 0; index < matchLength; index++) {
                    matrix[y][x++] = alive;
                }

            } while (matchElement)
        }

        return matrix;
    }
}

class Renderer {
    private static DEFAULT_CELL_WIDTH = 4;
    private static DEFAULT_CELL_HEIGHT = 4;
    private static DEFAULT_CELL_COLOR = "green";

    constructor(private readonly sheet: ExcelScript.Worksheet,
        private readonly cellWidth: number = Renderer.DEFAULT_CELL_WIDTH,
        private readonly cellHeight: number = Renderer.DEFAULT_CELL_HEIGHT,
        private readonly cellColor: string = Renderer.DEFAULT_CELL_COLOR) {
    }

    public initializeCanvas(width: number, height: number) {
        let address = `${Renderer.columnIndexToA1Address(0)}${1}:${Renderer.columnIndexToA1Address(width - 1)}${height}`
        let canvasRange = this.sheet.getRange(address);
        let format = canvasRange.getFormat();
        format.setColumnWidth(this.cellWidth);
        format.setRowHeight(this.cellHeight);
    }


    public renderEvolution(evolution: Evolution) {
        this.renderEvolvedCells(evolution.evolvedCells);
    }

    private renderEvolvedCells(evolvedCells: [number, number, boolean][]) {
        evolvedCells.forEach(evolvedCell => {
            const y = evolvedCell[0];
            const x = evolvedCell[1];
            const alive = evolvedCell[2];
            const fillFormat = this.sheet.getCell(y, x).getFormat().getFill();
            if (alive) {
                fillFormat.setColor(this.cellColor);
            } else {
                fillFormat.clear();
            }
        });
    }

    private static columnIndexToA1Address(column: number): string {
        let result = "";
        let current = column;
        while (current >= 0) {
            result = String.fromCharCode("A".charCodeAt(0) + (current % 26)) + result;
            current = Math.floor(current / 26) - 1;
        }
        return result;
    }

}

function sleep(milliseconds: number) {
    return new Promise(resolve => setTimeout(resolve, milliseconds));
}
