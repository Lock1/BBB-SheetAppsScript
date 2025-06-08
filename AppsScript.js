/** Brush's budget functional utilities for JS. v0.1 */
class Functionals {
    static identity() {
        return x => x;
    }

    static peek(consumer) {
        return x => { consumer(x); return x; };
    }

    /** Java's lazy stream / Haskell's list budget edition. Evaluation happen when terminal operation `collect()` is called. */
    static stream(arr, mapAccumulator=Functionals.identity()) {
        return {
            map: fn => Functionals.stream(arr, x => fn(mapAccumulator(x))),
            collect: () => arr.map(mapAccumulator),
        };
    }

    /** `IntStream` budget edition. See `Functionals.stream()`. */
    static intStream(start, end, mapAccumulator=Functionals.identity()) {
        const sink = consumer => { for (let i = start; i < end; ++i) consumer(mapAccumulator(i)); }
        return {
            map: fn => Functionals.intStream(start, end, x => fn(mapAccumulator(x))),
            collect: () => {
                const result = [];
                sink(x => result.push(x));
                return result;
            },
            sink,
        };
    }
}

/** It should be inside of `Functionals` namespace, but JS sucks */
class Pipe {
    /** Create a pipeline using provided `fn1` as the base pipe */
    static inlet(fn1, fnAccumulator=Functionals.identity()) {
        const composed = x => fn1(fnAccumulator(x));
        return {
            join: fn2 => Pipe.inlet(fn2, composed),
            outlet: fn2 => ({ compute: x => fn2(composed(x)) }), // JS doesn't need specialized API for sink
        };
    }

    /** Create a pipeline with the inlet connected to a source (side-effect data producer / constant function) */
    static source(fn1) {
        return Pipe.inlet(_ => fn1());
    }
}

class GoogleSheetUtils {
    static sheetRangeToLinearCellList(range) {
        return Functionals.intStream(1, range.getNumRows()+1)
            .map(i => Functionals.intStream(1, range.getNumColumns()+1).map(j => range.getCell(i, j)).collect())
            .collect()
            .flat();
    }

    static isCellInsideRange(range, cell) { // Ugh, sucks. No type => Hungarian notation again
        const row = cell.getRow();
        const col = cell.getColumn();
        return range.getRow() <= row && row <= range.getLastRow()
            && range.getColumn() <= col && col <= range.getLastColumn();
    }

    static isCellInsideFormatRange(format, cell) {
        return format.getRanges().some(range => GoogleSheetUtils.isCellInsideRange(range, cell));
    }
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Brush's Bargain Bin Scripts")
        .addSubMenu(ui.createMenu("Conditional Format")
            .addSubMenu(ui.createMenu("From selected range")
                .addItem("Enumerate categorical color scale", "categoricalFormatter")
                .addItem("Copy conditional format", "copyConditionalFormat")
                .addItem("Erase all in selection", "eraseConditionalFormat")
            )
        ).addToUi();
}

function categoricalFormatter() {
    function interpolator(percentage) {
        function lerp(start, end, t) { return start*(1-t) + end*t; }
        function integerLerp(start, end, t) { return Math.floor(lerp(start, end, t)); }
        function hexStringToTripletArray(str) {
            function hexBytes(str, start) { return str.substring(start, start+2); }
            return Functionals.intStream(0, 3).map(i => Number.parseInt(hexBytes(str, 2*i+1), 16)).collect();
        }
        const CONFIGURATION = {
            color : {
                mildOrangePurple: {
                    start: hexStringToTripletArray("#E59036"), // #E59036
                    mid:   hexStringToTripletArray("#76A5AF"), // #76A5AF
                    end:   hexStringToTripletArray("#D4A5BC"), // #D4A5BC
                },
                orangePurple: {
                    start: hexStringToTripletArray("#FF9900"), // #FF9900
                    mid:   hexStringToTripletArray("#45818E"), // #45818E
                    end:   hexStringToTripletArray("#A64D79"), // #A64D79
                },
                purpleOrangeBlue: {
                    start: hexStringToTripletArray("#4a86e8"), // #4a86e8
                    mid:   hexStringToTripletArray("#ff9900"), // #ff9900
                    end:   hexStringToTripletArray("#c27ba0"), // #c27ba0
                }
            },
            midpoint: 0.5,
        };
        const selectedColorScheme = CONFIGURATION.color.orangePurple;
        return Functionals.intStream(0, 3)
            .map(percentage < CONFIGURATION.midpoint
                ? i => integerLerp(selectedColorScheme.start[i], selectedColorScheme.mid[i], percentage/CONFIGURATION.midpoint)
                : i => integerLerp(selectedColorScheme.mid[i],   selectedColorScheme.end[i], (percentage-CONFIGURATION.midpoint)/(1-CONFIGURATION.midpoint))
            ).collect();
    }
    function tripletArrayToHexString(rgb) {
        const [r, g, b] = rgb.map(byte => byte.toString(16).padStart(2, "0"));
        return `#${r}${g}${b}`;
    }
    function keyColorRangeTupleToFormat([key, color, range]) {
        return SpreadsheetApp.newConditionalFormatRule()
            .whenTextEqualTo(key)
            .setBackground(tripletArrayToHexString(color))
            .setRanges([range])
            .build();
    }

    const selectedSheet = SpreadsheetApp.getActiveSheet();
    const selectedRange = selectedSheet.getSelection().getActiveRange();
    const formatListFromEnumerables = GoogleSheetUtils.sheetRangeToLinearCellList(selectedRange)
        .map(cell => cell.getValue())
        .filter((x, i, arr) => arr.indexOf(x) === i) // Naive duplicate filter
        .map((key, i, arr) => [key, interpolator(i/arr.length), selectedRange])
        .map(Functionals.peek(([key, color, _]) => Logger.log([key, tripletArrayToHexString(color)])))
        .map(keyColorRangeTupleToFormat);
    const prompt = `Selected Range: ${selectedRange.getA1Notation()}\n`
        + `Conditional format count: ${formatListFromEnumerables.length}\n`
        + `Apply categorical format?`;
    const ui = SpreadsheetApp.getUi();
    if (ui.alert(prompt, ui.ButtonSet.YES_NO) === ui.Button.YES) {
        Pipe.source(selectedSheet.getConditionalFormatRules)
            .join(originalFormatList => originalFormatList.concat(formatListFromEnumerables))
            .outlet(selectedSheet.setConditionalFormatRules)
            .compute();
    }
}

function copyConditionalFormat() {
    const selectedSheet = SpreadsheetApp.getActiveSheet();
    const selectedRange = selectedSheet.getSelection().getActiveRange();
    const cellList = GoogleSheetUtils.sheetRangeToLinearCellList(selectedRange);
    const selectedFormatList = selectedSheet.getConditionalFormatRules() // Filter out all format that doesn't contain cellList element
        .filter(format => cellList.some(cell => GoogleSheetUtils.isCellInsideFormatRange(format, cell)));

    const prompt = `Selected Range: ${selectedRange.getA1Notation()}\n`
        + `Conditional format count: ${selectedFormatList.length}\n`
        + `Copy?`;
    const ui = SpreadsheetApp.getUi();
    if (ui.alert(prompt, ui.ButtonSet.YES_NO) === ui.Button.YES) {
        const stringTargetSheet = ui.prompt("Type target sheet name:\n").getResponseText();
        const stringTargetRange = ui.prompt("Type target range using A1 notation\n").getResponseText();
        try {
            const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(stringTargetSheet);
            const targetRange = targetSheet.getRange(stringTargetRange);
            const clonedFormatList = selectedFormatList.map(format => format.copy().setRanges([targetRange]).build());
            Pipe.source(targetSheet.getConditionalFormatRules)
                .join(originalFormatList => originalFormatList.concat(clonedFormatList))
                .outlet(targetSheet.setConditionalFormatRules)
                .compute();
            ui.alert(`Format copied successfully (Sheet name: ${stringTargetSheet}, Range: ${stringTargetRange})`);
        } catch {
            ui.alert(`Error while trying to copy (Sheet name: ${stringTargetSheet}, Range: ${stringTargetRange})`);
        }
    }
}

function eraseConditionalFormat() {
    const selectedSheet = SpreadsheetApp.getActiveSheet();
    const selectedRange = selectedSheet.getSelection().getActiveRange();
    const cellList = GoogleSheetUtils.sheetRangeToLinearCellList(selectedRange);
    const ui = SpreadsheetApp.getUi();
    Pipe.source(selectedSheet.getConditionalFormatRules)
        .join(originalFormatList => {
            // Filter out all format that contain cellList element
            const filteredList = originalFormatList.filter(format => !cellList.some(cell => GoogleSheetUtils.isCellInsideFormatRange(format, cell)));
            const prompt = `Erase ${originalFormatList.length - filteredList.length} conditional format?\n`
                + `Note: This will completely erase it from the sheet, unlike "Clear Formatting" which only detach range`;
            return ui.alert(prompt, ui.ButtonSet.YES_NO) === ui.Button.YES ? filteredList : originalFormatList;
        }).outlet(selectedSheet.setConditionalFormatRules)
        .compute();
}

