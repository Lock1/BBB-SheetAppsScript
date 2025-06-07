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
}

function onOpen() {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu("Brush's Bargain Bin Scripts")
        .addSubMenu(
            ui.createMenu("Conditional Format")
                .addItem("Enumerate categorical color scale", "categoricalFormatter")
                .addItem("Copy", "copyFormat")
                .addItem("Erase all in selected range", "eraseFormat")
        ).addToUi();
}

function categoricalFormatter() {
    function interpolator(percentage) {
        function lerp(start, end, t) { return start*(1-t) + end*t; }
        function integerLerp(start, end, t) { return Math.floor(lerp(start, end, t)); }
        const CONFIGURATION = {
            color : {
                start: [0xE6, 0x90, 0x36],
                mid:   [0x76, 0xA5, 0xAF],
                end:   [0xD4, 0xA5, 0xBC],
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
    const enumerables = GoogleSheetUtils.sheetRangeToLinearCellList(selectedRange).map(cell => cell.getValue());
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

function eraseFormat() {
    function isCellInsideFormatRange(format, cell) {
        return format.getRanges().reduce(Functionals.anyReduce(range => GoogleSheetUtils.isCellInsideRange(range, cell)), false);
    }

    const selectedSheet = SpreadsheetApp.getActiveSheet();
    const selectedRange = selectedSheet.getSelection().getActiveRange();
    const cellList = GoogleSheetUtils.sheetRangeToLinearCellList(selectedRange);
    Functionals.attributeAccessorToMutationMap(
        selectedSheet.getConditionalFormatRules,
        selectedSheet.setConditionalFormatRules,
    ).map(
        originalFormatList => originalFormatList.filter(format => !cellList.reduce(
            Functionals.anyReduce(cell => isCellInsideFormatRange(format, cell)),
            false
        ))
    );
}
