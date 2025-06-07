class Functionals {
    static identity() {
        return x => x;
    }

    /** Lazy stream budget edition. Evaluation happen when terminal operation `collect()` is called. */
    static intStream(start, end, mapAccumulator=Functionals.identity()) {
        return {
            map: fn => Functionals.intStream(start, end, v => fn(mapAccumulator(v))),
            collect: () => {
                const result = [];
                for (let i = start; i < end; ++i) result.push(mapAccumulator(i));
                return result;
            },
        };
    }

    // Here's my original name: tranformTerribleOOPCargoCultGetSetIntoSanerMutationMap()
    /** Transform pair of getter-setter into pipelined functional programming map operation */
    static attributeAccessorToMutationMap(getter, setter) {
        return { map: mapper => setter(mapper(getter())) };
    }

    /** Use false as initial value */
    static anyReduce(predicate) {
        return (acc, x) => acc || predicate(x); 
    }
}

class GoogleSheetUtils {
    static sheetRangeToLinearCellList(range) {
        const result = [];
        for (let i = 1; i <= range.getNumRows(); ++i)
            for (let j = 1; j <= range.getNumColumns(); ++j)
                result.push(range.getCell(i, j));
        return result;
    }

    static isCellInsideRange(range, cell) { // Ugh, sucks. No type => Hungarian notation again
        const row = cell.getRow();
        const col = cell.getColumn();
        return range.getRow() <= row && row <= range.getLastRow()
            && range.getColumn() <= col && col <= range.getLastColumn();
    }

    static isCellInsideFormatRange(format, cell) {
        return format.getRanges().reduce(Functionals.anyReduce(range => GoogleSheetUtils.isCellInsideRange(range, cell)), false);
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
                        .addItem("Erase all in selection", "eraseFormat")
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
        return Functionals.intStream(0, 3).map(percentage < CONFIGURATION.midpoint
            ? i => integerLerp(CONFIGURATION.color.start[i], CONFIGURATION.color.mid[i], percentage/CONFIGURATION.midpoint)
            : i => integerLerp(CONFIGURATION.color.mid[i],   CONFIGURATION.color.end[i], (percentage-CONFIGURATION.midpoint)/(1-CONFIGURATION.midpoint))
        ).collect();
    }
    function tripletArrayToHexString(arr) {
        return `#${arr[0].toString(16)}${arr[1].toString(16)}${arr[2].toString(16)}`;
    }
    const selectedSheet = SpreadsheetApp.getActiveSheet();
    const selectedRange = selectedSheet.getSelection().getActiveRange();
    const enumerables = GoogleSheetUtils.sheetRangeToLinearCellList(selectedRange)
        .map(cell => cell.getValue())
        .filter((x, i, arr) => arr.indexOf(x) === i);
    const prompt = `Selected Range: ${selectedRange.getA1Notation()}\n`
        + `Conditional format count: ${enumerables.length}\n`
        + `Apply categorical format?`;
    const ui = SpreadsheetApp.getUi();
    if (ui.alert(prompt, ui.ButtonSet.YES_NO) === ui.Button.YES) {
        Functionals.attributeAccessorToMutationMap(
            selectedSheet.getConditionalFormatRules,
            selectedSheet.setConditionalFormatRules,
        ).map(originalFormatList =>
            originalFormatList.concat(
                enumerables.map((key, i) => [key, interpolator(i/enumerables.length)])
                    .map(pair => SpreadsheetApp.newConditionalFormatRule()
                            .whenTextEqualTo(pair[0])
                            .setBackground(tripletArrayToHexString(pair[1]))
                            .setRanges([selectedRange])
                            .build()
                        )
            )
        );
    }
}

function copyConditionalFormat() {
    const selectedSheet = SpreadsheetApp.getActiveSheet();
    const selectedRange = selectedSheet.getSelection().getActiveRange();
    const cellList = GoogleSheetUtils.sheetRangeToLinearCellList(selectedRange);
    const selectedFormatList = selectedSheet.getConditionalFormatRules() // Filter out all format that doesn't contain cellList element
        .filter(format => cellList.reduce(Functionals.anyReduce(cell => GoogleSheetUtils.isCellInsideFormatRange(format, cell)), false));

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
            Functionals.attributeAccessorToMutationMap(
                targetSheet.getConditionalFormatRules,
                targetSheet.setConditionalFormatRules,
            ).map(originalFormatList => originalFormatList.concat(clonedFormatList));
            ui.alert(`Format copied successfully (Sheet name: ${stringTargetSheet}, Range: ${stringTargetRange})`);
        } catch {
            ui.alert(`Error while trying to copy (Sheet name: ${stringTargetSheet}, Range: ${stringTargetRange})`);
        }
    }
}

function eraseFormat() {
    const selectedSheet = SpreadsheetApp.getActiveSheet();
    const selectedRange = selectedSheet.getSelection().getActiveRange();
    const cellList = GoogleSheetUtils.sheetRangeToLinearCellList(selectedRange);
    const ui = SpreadsheetApp.getUi();
    Functionals.attributeAccessorToMutationMap(
        selectedSheet.getConditionalFormatRules,
        selectedSheet.setConditionalFormatRules,
    ).map(
        originalFormatList => {
            const filteredList = originalFormatList.filter( // Filter out all format that contain cellList element
                format => !cellList.reduce(Functionals.anyReduce(cell => GoogleSheetUtils.isCellInsideFormatRange(format, cell)), false)
            );
            const prompt = `Erase ${originalFormatList.length - filteredList.length} conditional format?\n`
                + `Note: This will completely erase it from the sheet, unlike "Clear Formatting" which only detach range`;
            return ui.alert(prompt, ui.ButtonSet.YES_NO) === ui.Button.YES ? filteredList : originalFormatList;
        }
    );
}

