const pdfFile = document.getElementById("pdfFile");
const excelFile = document.getElementById("excelFile");

const extractBtn = document.getElementById("extractBtn");
const lookupBtn = document.getElementById("lookupBtn");
const exportBtn = document.getElementById("exportBtn");

const sheetNameInput = document.getElementById("sheetName");
const weldColumnInput = document.getElementById("weldColumn");
const startRowInput = document.getElementById("startRow");

const status = document.getElementById("status");
const result = document.getElementById("result");


/* =========================================================
   DATA
========================================================= */

let pdfRows = [];

let excelWorkbook = null;
let excelWorksheet = null;

let lookupResult = [];

let pdfFileName = "";
let excelFileName = "";


/* =========================================================
   1. DOC PDF
   PDF row structure:

   [Page, Text, X, Y]
========================================================= */

extractBtn.addEventListener("click", async function () {

    if (!pdfFile.files.length) {

        status.textContent =
            "Chua chon file PDF";

        return;
    }


    const file =
        pdfFile.files[0];


    pdfFileName =
        file.name;


    extractBtn.disabled = true;
    lookupBtn.disabled = true;
    exportBtn.disabled = true;


    result.innerHTML = "";


    status.textContent =
        "Dang doc PDF...";


    try {

        const arrayBuffer =
            await file.arrayBuffer();


        const pdf =
            await window.pdfjsLib.getDocument({
                data: arrayBuffer
            }).promise;


        pdfRows = [];


        for (
            let pageNumber = 1;
            pageNumber <= pdf.numPages;
            pageNumber++
        ) {

            status.textContent =
                "Dang doc trang " +
                pageNumber +
                "/" +
                pdf.numPages;


            const page =
                await pdf.getPage(
                    pageNumber
                );


            const textContent =
                await page.getTextContent();


            textContent.items.forEach(
                function (item) {

                    const text =
                        keyOf(item.str);


                    if (!text) {
                        return;
                    }


                    pdfRows.push([

                        pageNumber,

                        text,

                        item.transform[4],

                        item.transform[5]

                    ]);

                }
            );

        }


        status.textContent =
            "Da doc xong " +
            pdf.numPages +
            " trang. Tim thay " +
            pdfRows.length +
            " text.";


        updateLookupButton();

    }
    catch (error) {

        console.error(error);


        pdfRows = [];


        status.textContent =
            "Loi khi doc PDF: " +
            error.message;


        updateLookupButton();

    }
    finally {

        extractBtn.disabled = false;

    }

});


/* =========================================================
   2. DOC EXCEL
   Chi doc Workbook + Sheet.

   KHONG DOC HEADER.
   KHONG XAC DINH COT TU HEADER.

   Cot + dong bat dau se duoc xu ly
   khi bam "Tra cuu".
========================================================= */

excelFile.addEventListener(
    "change",
    async function () {

        if (!excelFile.files.length) {

            excelWorkbook = null;
            excelWorksheet = null;

            updateLookupButton();

            return;
        }


        const file =
            excelFile.files[0];


        excelFileName =
            file.name;


        try {

            status.textContent =
                "Dang doc Excel...";


            const arrayBuffer =
                await file.arrayBuffer();


            excelWorkbook =
                XLSX.read(
                    arrayBuffer,
                    {
                        type: "array"
                    }
                );


            /*
                Neu khong nhap Sheet
                thi dung Sheet dau tien.
            */

            let sheetName =
                sheetNameInput.value.trim();


            if (!sheetName) {

                sheetName =
                    excelWorkbook.SheetNames[0];


                sheetNameInput.value =
                    sheetName;

            }


            const matchedSheetName =
                excelWorkbook.SheetNames.find(
                    function (name) {

                        return name.trim() === sheetName
                            || name.trim().toLowerCase() === sheetName.toLowerCase();

                    }
                );


            if (!matchedSheetName) {

                throw new Error(
                    "Khong tim thay Sheet: " +
                    sheetName
                );

            }


            excelWorksheet =
                excelWorkbook.Sheets[
                    matchedSheetName
                ];


            status.textContent =
                "Da doc Excel. Sheet: " +
                sheetName +
                ". Hay khai bao cot va dong bat dau.";


            updateLookupButton();

        }
        catch (error) {

            console.error(error);


            excelWorkbook = null;
            excelWorksheet = null;


            status.textContent =
                "Loi khi doc Excel: " +
                error.message;


            updateLookupButton();

        }

    }
);


/* =========================================================
   3. TRA CUU
========================================================= */

lookupBtn.addEventListener(
    "click",
    function () {

        try {

            /*
                Kiem tra Excel
            */

            if (!excelWorksheet) {

                throw new Error(
                    "Chua doc file Excel."
                );

            }


            /*
                Lay cot moi han.

                Vi du:

                D -> 3
                AA -> 26
            */

            const sourceColumn =
                columnToIndex(
                    weldColumnInput.value.trim()
                );


            /*
                Dong bat dau tinh tu 1.

                Vi du:

                8 -> D8
            */

            const startRow =
                Number(
                    startRowInput.value
                );


            if (
                sourceColumn < 0
            ) {

                throw new Error(
                    "Cot ten moi han khong hop le."
                );

            }


            if (
                !Number.isInteger(startRow) ||
                startRow < 1
            ) {

                throw new Error(
                    "Dong bat dau phai la so nguyen >= 1."
                );

            }


            /*
                ------------------------------------------------
                DOC TRUC TIEP COT EXCEL
                ------------------------------------------------

                Neu:

                Cot = D
                Dong = 8

                thi doc:

                D8
                D9
                D10
                D11
                ...

                Khong quan tam Header.
            */

            const weldNames =
                readWeldNamesFromExcel(
                    excelWorksheet,
                    sourceColumn,
                    startRow
                );


            if (!weldNames.length) {

                throw new Error(
                    "Khong co ten moi han trong vung da chon."
                );

            }


            /*
                Chuyen thanh dang:

                [
                    ["FSW5/B1"],
                    ["FSW5/B2"],
                    ["FSW8"]
                ]

                de processWeldmap()
                xu ly.
            */

            const sourceRows =
                weldNames.map(
                    function (name) {

                        return [name];

                    }
                );


            /*
                TRA CUU PDF
            */

            lookupResult =
                processWeldmap({

                    sourceRows:
                        sourceRows,

                    pdfRows:
                        pdfRows,

                    sourceColumn: 0,

                    sourceStartRow: 1

                });


            /*
                Hien thi
            */

            showLookupResult(
                lookupResult
            );


            exportBtn.disabled =
                lookupResult.rows.length === 0;


            status.textContent =
                "Da doc " +
                lookupResult.totalCount +
                " ten moi han tu " +
                columnIndexToLetters(
                    sourceColumn
                ) +
                startRow +
                ". Tim thay " +
                lookupResult.matchedCount +
                " trong PDF.";

        }
        catch (error) {

            console.error(error);


            lookupResult = [];


            exportBtn.disabled =
                true;


            result.innerHTML = "";


            status.textContent =
                "Loi tra cuu: " +
                error.message;

        }

    }
);


exportBtn.addEventListener(
    "click",
    function () {

        try {

            if (
                !Array.isArray(lookupResult?.rows) ||
                lookupResult.rows.length === 0
            ) {

                throw new Error(
                    "Khong co du lieu de xuat Excel."
                );

            }


            const exportRows = [
                [
                    "Ten moi han",
                    "Trang",
                    "X",
                    "Y"
                ]
            ];


            lookupResult.rows.forEach(
                function (row) {

                    exportRows.push([
                        keyOf(row[0]),
                        keyOf(row[1]),
                        row[2] ?? "",
                        row[3] ?? ""
                    ]);

                }
            );


            const worksheet =
                XLSX.utils.aoa_to_sheet(
                    exportRows
                );


            const workbook =
                XLSX.utils.book_new();


            XLSX.utils.book_append_sheet(
                workbook,
                worksheet,
                "KetQua"
            );


            const fileName =
                (
                    pdfFileName
                        ? removeExtension(pdfFileName)
                        : "ket-qua-tra-cuu"
                ) +
                "_result.xlsx";


            if (
                typeof XLSX.writeFile === "function"
            ) {

                XLSX.writeFile(
                    workbook,
                    fileName
                );

            }
            else {

                const buffer =
                    XLSX.write(
                        workbook,
                        {
                            type: "array",
                            bookType: "xlsx"
                        }
                    );


                const blob =
                    new Blob(
                        [buffer],
                        {
                            type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        }
                    );


                const url =
                    URL.createObjectURL(blob);


                const link =
                    document.createElement("a");


                link.href = url;
                link.download = fileName;
                link.style.display = "none";
                link.target = "_blank";

                document.body.appendChild(link);
                link.click();
                link.remove();

                setTimeout(
                    function () {

                        URL.revokeObjectURL(url);

                    },
                    1000
                );

            }


            status.textContent =
                "Da xuat file Excel: " +
                fileName;

        }
        catch (error) {

            console.error(error);

            status.textContent =
                "Loi xuat Excel: " +
                error.message;

        }

    }
);


/* =========================================================
   4. DOC TRUC TIEP COT MOI HAN TU EXCEL
========================================================= */


function readWeldNamesFromExcel(
    worksheet,
    sourceColumn,
    startRow
) {

    const weldNames = [];

    /*
        Tim dong cuoi cung thuc te trong Sheet
        bang cach doc cac cell dang ton tai.
    */

    let lastRow = startRow - 1;


    for (const cellAddress of Object.keys(worksheet)) {

        /*
            Bo qua cac metadata cua Sheet:
            !ref
            !margins
            !merges
            ...
        */

        if (cellAddress.startsWith("!")) {
            continue;
        }


        const cell =
            XLSX.utils.decode_cell(
                cellAddress
            );


        /*
            cell.r = row index bat dau tu 0
        */

        if (cell.r >= startRow - 1) {

            if (cell.r > lastRow) {

                lastRow =
                    cell.r;

            }

        }

    }


    /*
        Khong co cell nao tu dong bat dau
    */

    if (
        lastRow < startRow - 1
    ) {

        return [];

    }


    /*
        Doc truc tiep:

        D8
        D9
        D10
        D11
        ...

        Neu sourceColumn = 3
        thi day la cot D.
    */

    for (
        let rowIndex = startRow - 1;
        rowIndex <= lastRow;
        rowIndex++
    ) {

        const cellAddress =
            XLSX.utils.encode_cell({

                r: rowIndex,

                c: sourceColumn

            });


        const cell =
            worksheet[cellAddress];


        if (!cell) {
            continue;
        }


        const value =
            cellTextOf(cell);


        if (!value) {
            continue;
        }


        weldNames.push(
            value
        );

    }


    return weldNames;

}

/* =========================================================
   5. PROCESS WELDMAP
   Logic tu weldmap-lookup.js

   Output:

   [Ten moi han, Trang, X, Y]
========================================================= */

function processWeldmap({

    sourceRows,

    pdfRows,

    sourceColumn = 0,

    sourceStartRow = 1,

    pdfKeyColumn = 1,

    pdfValueColumns = [0, 2, 3]

} = {}) {


    validateRows(
        sourceRows,
        "sourceRows"
    );


    validateRows(
        pdfRows,
        "pdfRows"
    );


    /*
        Excel sourceStartRow tinh tu 1.
    */

    const firstSourceIndex =
        Math.max(
            0,
            sourceStartRow - 1
        );


    const sourceData =
        sourceRows.slice(
            firstSourceIndex
        );


    /*
        Chi lay dong co ten moi han.
    */

    const dataRows =
        sourceData.filter(
            function (row) {

                return keyOf(
                    row?.[sourceColumn]
                ) !== "";

            }
        );


    if (!dataRows.length) {

        throw new Error(
            "Khong co ten moi han trong cot nguon."
        );

    }


    /*
        Tao Exact Lookup.
    */

    const duplicateKeys =
        new Set();


    const lookup =
        buildExactLookup(

            pdfRows,

            pdfKeyColumn,

            pdfValueColumns,

            duplicateKeys

        );


    const outputRows = [];

    let matchedCount = 0;


    /*
        Duyet tung ten moi han.
    */

    for (
        const sourceRow of dataRows
    ) {

        const sourceKey =
            keyOf(
                sourceRow[
                    sourceColumn
                ]
            );


        if (!sourceKey) {
            continue;
        }


        /*
            ----------------------------------------
            BUOC 1
            EXACT MATCH
            ----------------------------------------
        */

        let pdfRowIndex =
            lookup.get(
                sourceKey
            ) ?? -1;


        /*
            ----------------------------------------
            BUOC 2
            COMPOUND RULE
            ----------------------------------------
        */

        if (
            pdfRowIndex < 0
        ) {

            pdfRowIndex =
                findByCompoundRule(

                    sourceKey,

                    pdfRows,

                    pdfKeyColumn,

                    pdfValueColumns

                );

        }


        let page = "";
        let x = "";
        let y = "";


        /*
            Tim thay trong PDF.
        */

        if (
            pdfRowIndex >= 0
        ) {

            const pdfRow =
                pdfRows[
                    pdfRowIndex
                ];


            /*
                pdfValueColumns:

                [0] = Page
                [1] = X
                [2] = Y
            */

            page =
                pdfRow[
                    pdfValueColumns[0]
                ] ?? "";


            x =
                pdfRow[
                    pdfValueColumns[1]
                ] ?? "";


            y =
                pdfRow[
                    pdfValueColumns[2]
                ] ?? "";


            matchedCount++;

        }


        /*
            Output:

            Ten moi han
            Trang
            X
            Y

            Khong tim thay:
            Trang / X / Y = blank
        */

        outputRows.push([

            sourceKey,

            page,

            x,

            y

        ]);

    }


    return {

        rows:
            outputRows,

        matchedCount:
            matchedCount,

        totalCount:
            outputRows.length,

        duplicateKeys:
            [
                ...duplicateKeys
            ]

    };

}


/* =========================================================
   6. EXACT LOOKUP
========================================================= */

function buildExactLookup(

    pdfRows,

    keyColumn,

    valueColumns,

    duplicateKeys

) {

    const lookup =
        new Map();


    for (
        let rowIndex = 0;
        rowIndex < pdfRows.length;
        rowIndex++
    ) {

        const row =
            pdfRows[
                rowIndex
            ];


        if (!row) {
            continue;
        }


        const key =
            keyOf(
                row[keyColumn]
            );


        if (!key) {
            continue;
        }


        /*
            Key trung.
        */

        if (
            lookup.has(key)
        ) {

            duplicateKeys.add(
                key
            );


            const previousIndex =
                lookup.get(
                    key
                );


            const previousRow =
                pdfRows[
                    previousIndex
                ];


            /*
                Neu dong truoc rong
                nhung dong hien tai co du lieu
                thi dung dong hien tai.
            */

            if (

                isEmptyPdfRow(
                    previousRow,
                    valueColumns
                )

                &&

                !isEmptyPdfRow(
                    row,
                    valueColumns
                )

            ) {

                lookup.set(
                    key,
                    rowIndex
                );

            }

        }
        else {

            lookup.set(
                key,
                rowIndex
            );

        }

    }


    return lookup;

}


/* =========================================================
   7. COMPOUND RULE
========================================================= */

function findByCompoundRule(

    sourceKey,

    pdfRows,

    keyColumn,

    valueColumns

) {


    const sourceParts =
        splitCompoundKey(
            sourceKey
        );


    /*
        Khong phai compound key.
    */

    if (!sourceParts) {

        return -1;

    }


    const sourcePrefix =
        sourceParts.prefix;


    const sourceSuffix =
        sourceParts.suffix;


    /*
        ==============================================
        RULE 1

        Excel:

        FSW5/B1

        PDF cung co:

        FSW5/B1
        ==============================================
    */

    for (
        let rowIndex = 0;
        rowIndex < pdfRows.length;
        rowIndex++
    ) {

        const pdfKey =
            keyOf(
                pdfRows[
                    rowIndex
                ]?.[
                    keyColumn
                ]
            );


        if (
            sameKey(
                pdfKey,
                sourceKey
            )
        ) {

            return rowIndex;

        }

    }


    /*
        ==============================================
        RULE 2

        PDF:

        FSW5

        dong ke ben canh:

        B1-B10

        Excel:

        FSW5/B7

        => lay Page/X/Y cua FSW5
        ==============================================
    */

    for (
        let rowIndex = 0;
        rowIndex < pdfRows.length;
        rowIndex++
    ) {

        const currentKey =
            keyOf(
                pdfRows[
                    rowIndex
                ]?.[
                    keyColumn
                ]
            );


        if (
            !sameKey(
                currentKey,
                sourcePrefix
            )
        ) {

            continue;

        }


        /*
            Kiem tra dong tren
            va dong duoi.
        */

        for (
            const offset of [-1, 1]
        ) {

            const adjacentIndex =
                rowIndex + offset;


            if (

                adjacentIndex < 0

                ||

                adjacentIndex >=
                    pdfRows.length

            ) {

                continue;

            }


            const adjacentKey =
                keyOf(
                    pdfRows[
                        adjacentIndex
                    ]?.[
                        keyColumn
                    ]
                );


            const range =
                parseLetterNumberRange(
                    adjacentKey
                );


            const suffixParts =
                splitLetterNumber(
                    sourceSuffix
                );


            if (

                range

                &&

                suffixParts

                &&

                range.prefix.toUpperCase()
                    ===
                    suffixParts.prefix.toUpperCase()

                &&

                suffixParts.number >=
                    range.start

                &&

                suffixParts.number <=
                    range.end

            ) {

                return rowIndex;

            }

        }

    }


    return -1;

}


/* =========================================================
   8. SPLIT COMPOUND KEY
========================================================= */

function splitCompoundKey(value) {

    const text =
        keyOf(value);


    const slashIndex =
        text.lastIndexOf("/");


    const dashIndex =
        text.lastIndexOf("-");


    const separatorIndex =
        Math.max(
            slashIndex,
            dashIndex
        );


    if (
        separatorIndex <= 0
    ) {

        return null;

    }


    return {

        prefix:
            text
                .slice(
                    0,
                    separatorIndex
                )
                .trim(),


        suffix:
            text
                .slice(
                    separatorIndex + 1
                )
                .trim()

    };

}


/* =========================================================
   9. SPLIT LETTER + NUMBER
========================================================= */

function splitLetterNumber(value) {

    const text =
        keyOf(value);


    const match =
        text.match(
            /^([A-Za-z]+)\s*(\d+)$/
        );


    if (!match) {

        return null;

    }


    return {

        prefix:
            match[1],

        number:
            Number(
                match[2]
            )

    };

}


/* =========================================================
   10. PARSE RANGE
========================================================= */

function parseLetterNumberRange(value) {

    let text =
        keyOf(value);


    if (!text) {

        return null;

    }


    /*
        Vi du:

        (B1-B10)

        thanh:

        B1-B10
    */

    text =
        text.replace(
            /^\((.*)\)$/,
            "$1"
        );


    /*
        Ho tro:

        B1-B10
        B1~B10
        B1 TO B10
    */

    const match =
        text.match(

            /^([A-Za-z]+)\s*(\d+)\s*(?:~|-|TO)\s*(?:[A-Za-z]+\s*)?(\d+)$/i

        );


    if (match) {

        return {

            prefix:
                match[1],

            start:
                Number(
                    match[2]
                ),

            end:
                Number(
                    match[3]
                )

        };

    }


    /*
        Neu chi co:

        B5
    */

    const single =
        splitLetterNumber(
            text
        );


    if (single) {

        return {

            prefix:
                single.prefix,

            start:
                single.number,

            end:
                single.number

        };

    }


    return null;

}


/* =========================================================
   11. CHECK EMPTY PDF ROW
========================================================= */

function isEmptyPdfRow(
    row,
    valueColumns
) {

    if (!row) {

        return true;

    }


    return valueColumns.every(
        function (column) {

            return keyOf(
                row[column]
            ) === "";

        }
    );

}


/* =========================================================
   12. NORMALIZE KEY
========================================================= */

function keyOf(value) {

    return String(
        value ?? ""
    )

        /*
            NBSP -> space
        */

        .replace(
            /\u00A0/g,
            " "
        )

        /*
            CR / LF / TAB -> space
        */

        .replace(
            /[\r\n\t]/g,
            " "
        )

        .trim();

}


function cellTextOf(cell) {

    if (!cell) {

        return "";

    }


    const candidates = [
        cell.w,
        cell.v,
        cell.f,
        cell.r,
        cell.t
    ];


    for (const candidate of candidates) {

        const text =
            keyOf(candidate);


        if (text) {

            return text;

        }

    }


    return "";

}


/* =========================================================
   13. COMPARE KEY
========================================================= */

function sameKey(a, b) {

    return (

        keyOf(a).toUpperCase()

        ===

        keyOf(b).toUpperCase()

    );

}


/* =========================================================
   14. VALIDATE ROWS
========================================================= */

function validateRows(
    rows,
    name
) {

    if (
        !Array.isArray(rows)
    ) {

        throw new Error(
            name +
            " phai la mang."
        );

    }

}


/* =========================================================
   15. EXCEL COLUMN
========================================================= */

function columnToIndex(value) {

    const text =
        String(
            value ?? ""
        )
        .trim()
        .toUpperCase();


    if (!text) {

        return -1;

    }


    /*
        Cho phep:

        A
        B
        D
        AA

        Hoac:

        1
        2
        4
        27
    */

    if (
        /^\d+$/.test(text)
    ) {

        const number =
            Number(text);


        return number >= 1

            ? number - 1

            : -1;

    }


    if (
        !/^[A-Z]+$/.test(text)
    ) {

        return -1;

    }


    let result = 0;


    for (
        const char of text
    ) {

        result =
            result * 26
            +
            (
                char.charCodeAt(0)
                -
                64
            );

    }


    return result - 1;

}


/* =========================================================
   16. COLUMN INDEX -> LETTER
========================================================= */

function columnIndexToLetters(
    columnIndex
) {

    let result = "";

    let number =
        columnIndex + 1;


    while (
        number > 0
    ) {

        const remainder =
            (number - 1) % 26;


        result =
            String.fromCharCode(
                65 + remainder
            )
            +
            result;


        number =
            Math.floor(
                (number - 1) / 26
            );

    }


    return result;

}


/* =========================================================
   17. REMOVE FILE EXTENSION
========================================================= */

function removeExtension(
    fileName
) {

    return String(
        fileName ?? ""
    )
    .replace(
        /\.[^.]+$/,
        ""
    );

}


/* =========================================================
   18. UPDATE LOOKUP BUTTON
========================================================= */

function updateLookupButton() {

    lookupBtn.disabled =

        !excelWorksheet

        ||

        pdfRows.length === 0;

}


/* =========================================================
   19. HIEN THI KET QUA
========================================================= */

function showLookupResult(
    data
) {

    let html = `

        <table>

            <thead>

                <tr>

                    <th>Ten moi han</th>

                    <th>Trang</th>

                    <th>X</th>

                    <th>Y</th>

                </tr>

            </thead>

            <tbody>

    `;


    data.rows.forEach(
        function (row) {

            html += `

                <tr>

                    <td>
                        ${escapeHtml(
                            row[0]
                        )}
                    </td>

                    <td>
                        ${escapeHtml(
                            row[1]
                        )}
                    </td>

                    <td>
                        ${formatNumber(
                            row[2]
                        )}
                    </td>

                    <td>
                        ${formatNumber(
                            row[3]
                        )}
                    </td>

                </tr>

            `;

        }
    );


    html += `

            </tbody>

        </table>

    `;


    result.innerHTML =
        html;

}


/* =========================================================
   20. FORMAT NUMBER
========================================================= */

function formatNumber(value) {

    const number =
        Number(value);


    if (
        Number.isFinite(number)
    ) {

        return number.toFixed(2);

    }


    return escapeHtml(
        value
    );

}


/* =========================================================
   21. ESCAPE HTML
========================================================= */

function escapeHtml(value) {

    return String(
        value ?? ""
    )

    .replace(
        /&/g,
        "&amp;"
    )

    .replace(
        /</g,
        "&lt;"
    )

    .replace(
        />/g,
        "&gt;"
    )

    .replace(
        /"/g,
        "&quot;"
    )

    .replace(
        /'/g,
        "&#039;"
    );

}