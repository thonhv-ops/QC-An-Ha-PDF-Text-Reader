const pdfFile = document.getElementById("pdfFile");
const extractBtn = document.getElementById("extractBtn");
const exportBtn = document.getElementById("exportBtn");
const status = document.getElementById("status");
const result = document.getElementById("result");

let pdfTextData = [];
let currentFileName = "";


/* ================================
   EXTRACT PDF
================================ */

extractBtn.addEventListener("click", async function () {

    if (!pdfFile.files.length) {

        status.textContent = "Chua chon file PDF";

        return;
    }


    const file = pdfFile.files[0];

    currentFileName = file.name;

    status.textContent = "Dang doc PDF...";


    try {

        const arrayBuffer = await file.arrayBuffer();


        const pdf = await window.pdfjsLib.getDocument({
            data: arrayBuffer
        }).promise;


        pdfTextData = [];


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
                await pdf.getPage(pageNumber);


            const textContent =
                await page.getTextContent();


            textContent.items.forEach(function (item) {

                if (!item.str.trim()) {
                    return;
                }


                pdfTextData.push({

                    Page: pageNumber,

                    Text: item.str,

                    X: item.transform[4],

                    Y: item.transform[5],

                    Width: item.width,

                    Height: item.height

                });

            });

        }


        status.textContent =
            "Da doc xong " +
            pdf.numPages +
            " trang. Tim thay " +
            pdfTextData.length +
            " text.";


        exportBtn.disabled = false;


        showResult();


    }
    catch (error) {

        console.error(error);

        status.textContent =
            "Loi khi doc PDF: " +
            error.message;

    }

});


/* ================================
   DISPLAY RESULT
================================ */

function showResult() {

    let html = `

        <table>

            <thead>

                <tr>

                    <th>Page</th>

                    <th>Text</th>

                    <th>X</th>

                    <th>Y</th>

                    <th>Width</th>

                    <th>Height</th>

                </tr>

            </thead>

            <tbody>

    `;


    pdfTextData.forEach(function (item) {

        html += `

            <tr>

                <td>${item.Page}</td>

                <td>${item.Text}</td>

                <td>${item.X.toFixed(2)}</td>

                <td>${item.Y.toFixed(2)}</td>

                <td>${item.Width.toFixed(2)}</td>

                <td>${item.Height.toFixed(2)}</td>

            </tr>

        `;

    });


    html += `

            </tbody>

        </table>

    `;


    result.innerHTML = html;

}


/* ================================
   EXPORT EXCEL
================================ */

exportBtn.addEventListener("click", function () {

    if (!pdfTextData.length) {

        return;
    }


    const excelData = pdfTextData.map(function (item) {

        return {

            Page: item.Page,

            Text: item.Text,

            X: Number(item.X.toFixed(2)),

            Y: Number(item.Y.toFixed(2)),

            Width: Number(item.Width.toFixed(2)),

            Height: Number(item.Height.toFixed(2))

        };

    });


    const worksheet =
        XLSX.utils.json_to_sheet(excelData);


    const workbook =
        XLSX.utils.book_new();


    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        "PDF_Text"
    );


    const outputFileName =
        currentFileName.replace(
            /\.pdf$/i,
            ""
        ) + "_Text.xlsx";


    XLSX.writeFile(
        workbook,
        outputFileName
    );


    status.textContent =
        "Da xuat Excel: " +
        outputFileName;

});