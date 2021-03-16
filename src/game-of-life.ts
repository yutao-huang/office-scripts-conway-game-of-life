// The board dimension (number of columns/rows).
const BOARD_WIDTH = 120;
const BOARD_HEIGHT = 60;

// The dimension of each cell.
const DEFAULT_CELL_WIDTH = 8;
const DEFAULT_CELL_HEIGHT = 8;

// The cell color
const DEFAULT_CELL_COLOR = "green";

// How many generations you want to evolve.
const NUMBER_OF_GENERATIONS = 120;

// Refer to https://copy.sh/life/examples for more sample patterns.
const PATTERN_URL = "https://copy.sh/life/examples/glider.rle";

async function main(workbook: ExcelScript.Workbook): Promise<void> {
    let sheet = workbook.addWorksheet();
    sheet.activate();

    const pattern = await Pattern.fromUrl(PATTERN_URL);
    console.log(`Pattern: ${pattern.name}, ${pattern.width} x ${pattern.height}; Rule: ${pattern.rule.identifier}, ${pattern.rule.name}`);
    if (!pattern.rule || pattern.rule.name === "Unsupported") {
        console.log(`The rule '${pattern.rule.identifier}' used by this pattern is not supported yet. Please pick another one.`);
        return;
    }

    const game = new Game(BOARD_WIDTH, BOARD_HEIGHT, pattern);
    const renderer = new Renderer(sheet);

    renderer.initializeCanvas(game.width, game.height);
    renderer.renderEvolution(game.getInitialEvolution());

    console.log("Evolving...");

    // Rendering might fail if the interval is too small. Normally it'd be fine if >= 500 milliseconds.
    const RENDER_INTERVAL_MILLISECONDS = 500;

    for (var generation = 1; generation < NUMBER_OF_GENERATIONS; generation++) {
        await sleep(RENDER_INTERVAL_MILLISECONDS);

        let evolution = game.evolveOneGeneration();
        if (!evolution.hasEvolved) {
            if (game.hasLife) {
                console.log(`Generation #${generation} has become still life.`);
            } else {
                console.log(`Unfortunately Generation #${generation} has become extinct.`);
            }
            break;
        }
        renderer.renderEvolution(evolution);
    }
}

interface Grid {
    width: number;
    height: number;
    matrix: boolean[][];
};

class Game implements Grid {
    readonly matrix: boolean[][];

    constructor(public readonly width: number, public readonly height: number, public readonly initialPattern: Pattern) {
        this.matrix = new Array(height).fill(false).map(() => new Array(width).fill(false));
    }

    get hasLife(): boolean {
        return this.matrix.some(row => row.some(isCellAlive => isCellAlive));
    }

    getInitialEvolution(): Evolution {
        let evolution = new Evolution();

        let patternX = Math.floor((this.width - this.initialPattern.width) / 2);
        let patternY = Math.floor((this.height - this.initialPattern.height) / 2);
        for (var y = 0; y < this.initialPattern.height; y++) {
            for (var x = 0; x < this.initialPattern.width; x++) {
                this.matrix[y + patternY][x + patternX] = this.initialPattern.matrix[y][x];
                if (this.initialPattern.matrix[y][x]) {
                    evolution.evolveCell(y + patternY, x + patternX, true);
                }
            }
        }

        return evolution;
    }

    evolveOneGeneration(): Evolution {
        let evolution = new Evolution();

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
        const currentlyAlive = this.initialPattern.rule.isCellAlive(previouslyAlive, neighbors);
        if (previouslyAlive !== currentlyAlive) {
            evolution.evolveCell(y, x, currentlyAlive);
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

interface Rule {
    identifier: string;
    name: string;
    isCellAlive(previouslyAlive: boolean, numberOfNeighbors: number): boolean;
}

class RuleFactory {
    buildRule(identifier: string): Rule {
        switch (identifier.toUpperCase()) {
            case "B3/S23":
                return new ConwayLifeRule;
            default:
                break;
        }
    }
}
class ConwayLifeRule implements Rule {
    readonly identifier = "B3/S23";
    readonly name = "Conway's Game of Life";
    
    isCellAlive(previouslyAlive: boolean, numberOfNeighbors: number): boolean {
        switch (true) {
            case (previouslyAlive && (numberOfNeighbors < 2 || numberOfNeighbors > 3)): return false;
            case (previouslyAlive && (numberOfNeighbors === 2 || numberOfNeighbors === 3)): return true;
            case (!previouslyAlive && numberOfNeighbors === 3): return true;
            default: return false;
        }
    }
}

class MoveRule implements Rule {
    readonly identifier = "245/368";
    readonly name = "Move (or Morley)";
    
    isCellAlive(previouslyAlive: boolean, numberOfNeighbors: number): boolean {
        switch (true) {
            case (previouslyAlive && [2, 4, 5].includes(numberOfNeighbors)): return true;
            case (!previouslyAlive && [3, 6, 8].includes(numberOfNeighbors)): return true;
            default: return false;
        }
    }
}

class HighLifeRule implements Rule {
    readonly identifier = "23/36";
    readonly name = "HighLife";
    
    isCellAlive(previouslyAlive: boolean, numberOfNeighbors: number): boolean {
        switch (true) {
            case (previouslyAlive && [2, 3].includes(numberOfNeighbors)): return true;
            case (!previouslyAlive && [3, 6].includes(numberOfNeighbors)): return true;
            default: return false;
        }
    }
}

class TwoByTwoRule implements Rule {
    readonly identifier = "125/36";
    readonly name = "2x2";
    isCellAlive(previouslyAlive: boolean, numberOfNeighbors: number): boolean {
        switch (true) {
            case (previouslyAlive && [1, 2, 5].includes(numberOfNeighbors)): return true;
            case (!previouslyAlive && [3, 6].includes(numberOfNeighbors)): return true;
            default: return false;
        }
    }
}

class MazeRule implements Rule {
    identifier = "12345/3";
    name: "Maze";
    isCellAlive(previouslyAlive: boolean, numberOfNeighbors: number): boolean {
        switch (true) {
            case (previouslyAlive && [1, 2, 3, 4, 5].includes(numberOfNeighbors)): return true;
            case (!previouslyAlive && numberOfNeighbors === 3): return true;
            default: return false;
        }
    }
}

class LifeWithoutDeathRule implements Rule {
    identifier = "b3/s012345678";
    name = "Life without death";
    isCellAlive(previouslyAlive: boolean, numberOfNeighbors: number): boolean {
        switch (true) {
            case (previouslyAlive): return true;
            case (!previouslyAlive && numberOfNeighbors === 3): return true;
            default: return false;
        }
    }
}

class UnsupportedRule implements Rule {
    readonly name = "Unsupported";

    constructor(readonly identifier: string) {
    }

    isCellAlive(previouslyAlive: boolean, numberOfNeighbors: number): boolean {
        throw new Error("Method not implemented.");
    }
}

class Pattern implements Grid {
    width: number;
    height: number;
    matrix: boolean[][];
    name: string;
    rule: Rule;

    private static readonly supportedRules = [
        new ConwayLifeRule,
        new HighLifeRule,
        new MoveRule,
        new TwoByTwoRule,
        new MazeRule,
        new LifeWithoutDeathRule
    ]

    static async fromUrl(url: string): Promise<Pattern> {
        let fetchResult = await fetch(`https://sofetch.glitch.me/${encodeURI(url)}`);
        let patternFileContent = await fetchResult.text();
        let lines = patternFileContent.split("\n");
        let pattern: Pattern = { width: 0, height: 0, matrix: null, name: "", rule: null };
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
                pattern.rule = Pattern.getRule(matchGroups.rule);
            } else {
                patternString += line;
            }
        });

        pattern.matrix = Pattern.parsePattern(patternString.replace("!", ""), pattern.width, pattern.height);

        return pattern;
    }

    private constructor() {
    }

    private static getRule(identifier: string): Rule {
        return Pattern.supportedRules.find(rule => rule.identifier.toUpperCase() === identifier.toUpperCase()) ?? new UnsupportedRule(identifier);
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

class Evolution {
    readonly evolvedCells: [number, number, boolean][] = new Array<[number, number, boolean]>();

    get hasEvolved(): boolean {
        return this.evolvedCells.length > 0;
    }

    evolveCell(y: number, x: number, alive: boolean): void {
        this.evolvedCells.push([y, x, alive]);
    }
}

class Renderer {
    constructor(private readonly sheet: ExcelScript.Worksheet,
        private readonly cellWidth: number = DEFAULT_CELL_WIDTH,
        private readonly cellHeight: number = DEFAULT_CELL_HEIGHT,
        private readonly cellColor: string = DEFAULT_CELL_COLOR) {
    }

    initializeCanvas(width: number, height: number) {
        let address = `${Renderer.columnIndexToA1Address(0)}${1}:${Renderer.columnIndexToA1Address(width - 1)}${height}`
        let canvasRange = this.sheet.getRange(address);
        let format = canvasRange.getFormat();
        format.setColumnWidth(this.cellWidth);
        format.setRowHeight(this.cellHeight);
    }


    renderEvolution(evolution: Evolution) {
        evolution.evolvedCells.forEach(evolvedCell => {
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
