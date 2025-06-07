/** Brush's budget functional utilities for JS. v0.1 */
class Functionals {
    static identity() {
        return x => x;
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
        return {
            map: fn => Functionals.intStream(start, end, x => fn(mapAccumulator(x))),
            collect: () => {
                const result = [];
                for (let i = start; i < end; ++i) result.push(mapAccumulator(i));
                return result;
            },
        };
    }

    // Here's my original name: transformTerribleOOPCargoCultGetSetIntoSanerMutationMap()
    /** Transform pair of getter-setter into pipelined functional programming map operation */
    static attributeAccessorToMutationMap(getter, setter) {
        return { map: mapper => setter(mapper(getter())) };
    }
}

/** It's Functionals nested namespace, but JS sucks */
class Pipe {
    static input(fn1, fnAccumulator=Functionals.identity()) {
        return {
            connect: fn2 => Pipe.input(fn2, x => fn1(fnAccumulator(x))),
            output: fn2 => ({ compute: x => fn2(fn1(fnAccumulator(x))) }),
        };
    }

    static source(fn1) {
        return Pipe.input(_ => fn1());
    }
}

class GoogleSheetUtils {
    static sheetRangeToLinearCellList(range) {
        return Functionals.intStream(1, range.getNumRows()+1)
            .map(i => Functionals.intStream(1, range.getNumColumns()+1).map(j => range.getCell(i, j)).collect()
        ).collect()
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
        .addSubMenu(
            ui.createMenu("Conditional Format")
                .addSubMenu(
                    ui.createMenu("From selected range")
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
        const CONFIGURATION = {
            color : {
                start: [0xE6, 0x90, 0x36], // Orange-ish
                mid:   [0x76, 0xA5, 0xAF], // Dark green-bluish
                end:   [0xD4, 0xA5, 0xBC], // Mild purple
            },
            midpoint: 0.5,
        };
        return Functionals.intStream(0, 3)
            .map(percentage < CONFIGURATION.midpoint
                ? i => integerLerp(CONFIGURATION.color.start[i], CONFIGURATION.color.mid[i], percentage/CONFIGURATION.midpoint)
                : i => integerLerp(CONFIGURATION.color.mid[i],   CONFIGURATION.color.end[i], (percentage-CONFIGURATION.midpoint)/(1-CONFIGURATION.midpoint))
            ).collect();
    }
    function tripletArrayToHexString([r, g, b]) {
        return `#${r.toString(16)}${g.toString(16)}${b.toString(16)}`;
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
        .map(keyColorRangeTupleToFormat);
    const prompt = `Selected Range: ${selectedRange.getA1Notation()}\n`
        + `Conditional format count: ${formatListFromEnumerables.length}\n`
        + `Apply categorical format?`;
    const ui = SpreadsheetApp.getUi();
    if (ui.alert(prompt, ui.ButtonSet.YES_NO) === ui.Button.YES) {
        Pipe.source(selectedSheet.getConditionalFormatRules)
            .connect(originalFormatList => originalFormatList.concat(formatListFromEnumerables))
            .output(selectedSheet.setConditionalFormatRules)
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
                .connect(originalFormatList => originalFormatList.concat(clonedFormatList))
                .output(targetSheet.setConditionalFormatRules)
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
        .connect(originalFormatList => {
            // Filter out all format that contain cellList element
            const filteredList = originalFormatList.filter(format => !cellList.some(cell => GoogleSheetUtils.isCellInsideFormatRange(format, cell)));
            const prompt = `Erase ${originalFormatList.length - filteredList.length} conditional format?\n`
                + `Note: This will completely erase it from the sheet, unlike "Clear Formatting" which only detach range`;
            return ui.alert(prompt, ui.ButtonSet.YES_NO) === ui.Button.YES ? filteredList : originalFormatList;
        }).output(selectedSheet.setConditionalFormatRules)
        .compute();
}

