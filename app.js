/* =====================================================
   FISH TRADE INTERACTIVE DASHBOARD
   D3.js + SheetJS + Leaflet
===================================================== */


/* =====================================================
   GLOBAL VARIABLES
===================================================== */

let rawData = [];

let filteredData = [];

let currentPage = 1;

const ROWS_PER_PAGE = 10;

let map = null;

let mapMarkers = [];

let tooltip = null;


/* =====================================================
   MONTH NAMES
===================================================== */

const monthNames = {

    1: "มกราคม",
    2: "กุมภาพันธ์",
    3: "มีนาคม",
    4: "เมษายน",
    5: "พฤษภาคม",
    6: "มิถุนายน",
    7: "กรกฎาคม",
    8: "สิงหาคม",
    9: "กันยายน",
    10: "ตุลาคม",
    11: "พฤศจิกายน",
    12: "ธันวาคม"

};


/* =====================================================
   PASTEL COLORS
===================================================== */

const pastelColors = [

    "#fd9097",
    "#f8c891",
    "#f8f862",
    "#6ef88c",
    "#74bef7",
    "#cdb4db",
    "#ffc8dd",
    "#b8e0d2",
    "#f6bdc8",
    "#bde0fe"

];


/* =====================================================
   COUNTRY COORDINATES
===================================================== */

const countryCoordinates = {

    "ไทย": [15.8700, 100.9925],

    "จีน": [35.8617, 104.1954],

    "ญี่ปุ่น": [36.2048, 138.2529],

    "เกาหลีใต้": [35.9078, 127.7669],

    "ไต้หวัน": [23.6978, 120.9605],

    "ฮ่องกง": [22.3193, 114.1694],

    "มาเลเซีย": [4.2105, 101.9758],

    "สิงคโปร์": [1.3521, 103.8198],

    "อินโดนีเซีย": [-0.7893, 113.9213],

    "เวียดนาม": [14.0583, 108.2772],

    "ฟิลิปปินส์": [12.8797, 121.7740],

    "อินเดีย": [20.5937, 78.9629],

    "มัลดีฟส์": [3.2028, 73.2207],

    "ออสเตรเลีย": [-25.2744, 133.7751],

    "สหรัฐอเมริกา": [37.0902, -95.7129],

    "แคนาดา": [56.1304, -106.3468],

    "สหราชอาณาจักร": [55.3781, -3.4360],

    "ฝรั่งเศส": [46.2276, 2.2137],

    "เยอรมนี": [51.1657, 10.4515],

    "เนเธอร์แลนด์": [52.1326, 5.2913],

    "สเปน": [40.4637, -3.7492],

    "อิตาลี": [41.8719, 12.5674],

    "กาตาร์": [25.3548, 51.1839],

    "สหรัฐอาหรับเอมิเรตส์": [23.4241, 53.8478]

};


/* =====================================================
   INITIALIZATION
===================================================== */

document.addEventListener(
    "DOMContentLoaded",
    () => {

        createTooltip();

        setupEventListeners();

        loadExcel();

    }
);


/* =====================================================
   TOOLTIP
===================================================== */

function createTooltip() {

    tooltip =
        d3.select("body")

            .append("div")

            .attr(
                "class",
                "chart-tooltip"
            );

}


/* =====================================================
   LOAD EXCEL
===================================================== */

async function loadExcel() {

    try {

        setStatus(
            "กำลังโหลดข้อมูล..."
        );


        /*
            IMPORTANT

            ไฟล์ Excel ต้องชื่อ data.xlsx
            และอยู่โฟลเดอร์เดียวกับ index.html
        */

        const response =
            await fetch(
                "data.xlsx"
            );


        if (!response.ok) {

            throw new Error(
                `ไม่พบไฟล์ data.xlsx (${response.status})`
            );

        }


        const arrayBuffer =
            await response.arrayBuffer();


        const workbook =
            XLSX.read(
                arrayBuffer,
                {
                    type: "array"
                }
            );


        /*
            ใช้ Sheet จริงจากไฟล์
            ซึ่งตรวจสอบแล้วชื่อ Cleaned_Data
        */

        const sheetName =
            workbook.SheetNames.includes(
                "Cleaned_Data"
            )
                ? "Cleaned_Data"
                : workbook.SheetNames[0];


        const worksheet =
            workbook.Sheets[
                sheetName
            ];


        const excelData =
            XLSX.utils.sheet_to_json(
                worksheet,
                {
                    defval: ""
                }
            );


        if (
            !excelData ||
            excelData.length === 0
        ) {

            throw new Error(
                "ไม่พบข้อมูลในไฟล์ Excel"
            );

        }


        rawData =
            excelData.map(
                normalizeRow
            );


        filteredData =
            [...rawData];


        setStatus(
            `โหลดข้อมูลสำเร็จ ${rawData.length.toLocaleString()} รายการ`
        );


        initializeFilters();

        initializeMap();

        updateDashboard();


    }
    catch (error) {

        console.error(
            "Excel Error:",
            error
        );


        setStatus(
            "โหลดข้อมูลไม่สำเร็จ"
        );


        showError(
            `เกิดข้อผิดพลาด: ${error.message}<br><br>
             ตรวจสอบว่าไฟล์ชื่อ <b>data.xlsx</b>
             และอยู่โฟลเดอร์เดียวกับ index.html`
        );

    }

}


/* =====================================================
   NORMALIZE ROW
===================================================== */

function normalizeRow(row) {

    return {

        year:
            toNumber(
                row.year
            ),

        month:
            toNumber(
                row.month
            ),

        heading11:
            String(
                row.heading11 ?? ""
            ),

        countryID:
            String(
                row.countryID ?? ""
            ),

        country:
            cleanText(
                row.countryNameTH,
                "ไม่ระบุ"
            ),

        weight:
            toNumber(
                row.weight
            ),

        quantity:
            toNumber(
                row.quantity
            ),

        price:
            toNumber(
                row.price
            ),

        tradeflow:
            cleanText(
                row.tradeflow,
                "ไม่ระบุ"
            ),

        productEN:
            cleanText(
                row.productDetailEN,
                "ไม่ระบุ"
            ),

        fish:
            cleanText(
                row.fishName,
                "ไม่ระบุ"
            ),

        productTH:
            cleanText(
                row.productDetailTH,
                "ไม่ระบุ"
            )

    };

}


/* =====================================================
   DATA HELPERS
===================================================== */

function toNumber(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return 0;

    }


    const number =
        Number(
            String(value)
                .replace(/,/g, "")
                .trim()
        );


    return Number.isFinite(number)
        ? number
        : 0;

}


function cleanText(
    value,
    fallback = ""
) {

    const text =
        String(
            value ?? ""
        ).trim();


    return text || fallback;

}


/* =====================================================
   STATUS
===================================================== */

function setStatus(message) {

    const element =
        document.getElementById(
            "dataStatus"
        );


    if (element) {

        element.textContent =
            message;

    }

}


/* =====================================================
   ERROR MESSAGE
===================================================== */

function showError(message) {

    const container =
        document.getElementById(
            "barChart"
        );


    container.innerHTML = `

        <div class="error-box">

            <strong>⚠️ ไม่สามารถโหลดข้อมูลได้</strong>

            <p>${message}</p>

        </div>

    `;

}


/* =====================================================
   EVENT LISTENERS
===================================================== */

function setupEventListeners() {

    document
        .getElementById("yearFilter")
        .addEventListener(
            "change",
            applyFilters
        );


    document
        .getElementById("monthFilter")
        .addEventListener(
            "change",
            applyFilters
        );


    document
        .getElementById("countryFilter")
        .addEventListener(
            "change",
            applyFilters
        );


    document
        .getElementById("tradeFilter")
        .addEventListener(
            "change",
            applyFilters
        );


    document
        .getElementById("resetButton")
        .addEventListener(
            "click",
            resetFilters
        );


    document
        .getElementById("searchInput")
        .addEventListener(
            "input",
            () => {

                currentPage = 1;

                updateTable();

            }
        );


    document
        .getElementById("prevPage")
        .addEventListener(
            "click",
            () => {

                if (
                    currentPage > 1
                ) {

                    currentPage--;

                    updateTable();

                }

            }
        );


    document
        .getElementById("nextPage")
        .addEventListener(
            "click",
            () => {

                const searchData =
                    getSearchData();


                const totalPages =
                    Math.max(
                        1,
                        Math.ceil(
                            searchData.length /
                            ROWS_PER_PAGE
                        )
                    );


                if (
                    currentPage <
                    totalPages
                ) {

                    currentPage++;

                    updateTable();

                }

            }
        );

}


/* =====================================================
   FILTER OPTIONS
===================================================== */

function initializeFilters() {

    const yearSelect =
        document.getElementById(
            "yearFilter"
        );


    const countrySelect =
        document.getElementById(
            "countryFilter"
        );


    const tradeSelect =
        document.getElementById(
            "tradeFilter"
        );


    /* Clear old options */

    yearSelect.innerHTML =
        `<option value="all">
            ทุกปี
        </option>`;


    countrySelect.innerHTML =
        `<option value="all">
            ทุกประเทศ
        </option>`;


    tradeSelect.innerHTML =
        `<option value="all">
            ทั้งหมด
        </option>`;


    /* Years */

    const years =
        [...new Set(
            rawData
                .map(
                    d => d.year
                )
                .filter(
                    d => d !== 0
                )
        )]
        .sort(
            (a,b) => a - b
        );


    years.forEach(
        year => {

            const option =
                document.createElement(
                    "option"
                );

            option.value =
                year;

            option.textContent =
                year;

            yearSelect.appendChild(
                option
            );

        }
    );


    /* Countries */

    const countries =
        [...new Set(
            rawData
                .map(
                    d => d.country
                )
                .filter(
                    Boolean
                )
        )]
        .sort(
            (a,b) =>
                a.localeCompare(
                    b,
                    "th"
                )
        );


    countries.forEach(
        country => {

            const option =
                document.createElement(
                    "option"
                );

            option.value =
                country;

            option.textContent =
                country;

            countrySelect.appendChild(
                option
            );

        }
    );


    /* Trade Flow */

    const tradeFlows =
        [...new Set(
            rawData
                .map(
                    d => d.tradeflow
                )
                .filter(
                    Boolean
                )
        )]
        .sort();


    tradeFlows.forEach(
        flow => {

            const option =
                document.createElement(
                    "option"
                );

            option.value =
                flow;

            option.textContent =
                `Trade Flow ${flow}`;

            tradeSelect.appendChild(
                option
            );

        }
    );

}


/* =====================================================
   APPLY FILTER
===================================================== */

function applyFilters() {

    const year =
        document.getElementById(
            "yearFilter"
        ).value;


    const month =
        document.getElementById(
            "monthFilter"
        ).value;


    const country =
        document.getElementById(
            "countryFilter"
        ).value;


    const trade =
        document.getElementById(
            "tradeFilter"
        ).value;


    filteredData =
        rawData.filter(
            d => {

                const yearOK =
                    year === "all" ||
                    String(d.year) === year;


                const monthOK =
                    month === "all" ||
                    String(d.month) === month;


                const countryOK =
                    country === "all" ||
                    d.country === country;


                const tradeOK =
                    trade === "all" ||
                    d.tradeflow === trade;


                return (
                    yearOK &&
                    monthOK &&
                    countryOK &&
                    tradeOK
                );

            }
        );


    currentPage = 1;

    updateDashboard();

}


/* =====================================================
   RESET
===================================================== */

function resetFilters() {

    document.getElementById(
        "yearFilter"
    ).value = "all";


    document.getElementById(
        "monthFilter"
    ).value = "all";


    document.getElementById(
        "countryFilter"
    ).value = "all";


    document.getElementById(
        "tradeFilter"
    ).value = "all";


    document.getElementById(
        "searchInput"
    ).value = "";


    filteredData =
        [...rawData];


    currentPage = 1;

    updateDashboard();

}


/* =====================================================
   UPDATE DASHBOARD
===================================================== */

function updateDashboard() {

    updateKPI();

    drawBarChart();

    drawDonutChart();

    drawLineChart();

    updateMap();

    updateTable();

}


/* =====================================================
   KPI
===================================================== */

function updateKPI() {

    const totalRecords =
        filteredData.length;


    const totalWeight =
        d3.sum(
            filteredData,
            d => d.weight
        );


    const totalQuantity =
        d3.sum(
            filteredData,
            d => d.quantity
        );


    const totalPrice =
        d3.sum(
            filteredData,
            d => d.price
        );


    animateNumber(
        "totalRecords",
        totalRecords
    );


    animateNumber(
        "totalWeight",
        totalWeight
    );


    animateNumber(
        "totalQuantity",
        totalQuantity
    );


    animateNumber(
        "totalPrice",
        totalPrice
    );

}


/* =====================================================
   KPI ANIMATION
===================================================== */

function animateNumber(
    elementId,
    value
) {

    const element =
        document.getElementById(
            elementId
        );


    if (!element) return;


    const oldValue =
        Number(
            element.dataset.value || 0
        );


    const startTime =
        performance.now();


    const duration =
        650;


    function animate(time) {

        const progress =
            Math.min(
                1,
                (time - startTime) /
                duration
            );


        const eased =
            1 -
            Math.pow(
                1 - progress,
                3
            );


        const current =
            oldValue +
            (value - oldValue) *
            eased;


        element.textContent =
            Math.round(
                current
            ).toLocaleString();


        if (
            progress < 1
        ) {

            requestAnimationFrame(
                animate
            );

        }

    }


    element.dataset.value =
        value;


    requestAnimationFrame(
        animate
    );

}


/* =====================================================
   BAR CHART
===================================================== */

function drawBarChart() {

    const container =
        document.getElementById(
            "barChart"
        );


    container.innerHTML = "";


    if (
        filteredData.length === 0
    ) {

        showNoData(
            container
        );

        return;

    }


    const grouped =
        d3.rollups(

            filteredData,

            values =>
                d3.sum(
                    values,
                    d => d.price
                ),

            d => d.country

        );


    const data =
        grouped

            .map(
                ([country,value]) => ({
                    country,
                    value
                })
            )

            .sort(
                (a,b) =>
                    d3.descending(
                        a.value,
                        b.value
                    )
            )

            .slice(
                0,
                10
            );


    const width =
        Math.max(
            container.clientWidth,
            500
        );


    const height =
        380;


    const margin = {

        top: 25,

        right: 25,

        bottom: 85,

        left: 75

    };


    const svg =
        d3.select(container)

            .append("svg")

            .attr(
                "viewBox",
                `0 0 ${width} ${height}`
            );


    const x =
        d3.scaleBand()

            .domain(
                data.map(
                    d => d.country
                )
            )

            .range([
                margin.left,
                width - margin.right
            ])

            .padding(
                0.22
            );


    const maxValue =
        d3.max(
            data,
            d => d.value
        ) || 1;


    const y =
        d3.scaleLinear()

            .domain([
                0,
                maxValue
            ])

            .nice()

            .range([
                height - margin.bottom,
                margin.top
            ]);


    /* Grid */

    svg.append("g")

        .attr(
            "class",
            "grid"
        )

        .attr(
            "transform",
            `translate(${margin.left},0)`
        )

        .call(
            d3.axisLeft(y)
                .ticks(5)
                .tickSize(
                    -(width -
                        margin.left -
                        margin.right)
                )
                .tickFormat("")
        )

        .selectAll("line")

        .attr(
            "class",
            "grid-line"
        );


    /* Bars */

    const bars =
        svg.selectAll(
            ".bar"
        )

        .data(data)

        .enter()

        .append("rect")

        .attr(
            "class",
            "bar"
        )

        .attr(
            "x",
            d => x(d.country)
        )

        .attr(
            "width",
            x.bandwidth()
        )

        .attr(
            "y",
            height - margin.bottom
        )

        .attr(
            "height",
            0
        )

        .attr(
            "rx",
            7
        )

        .attr(
            "fill",
            (d,i) =>
                pastelColors[
                    i % pastelColors.length
                ]
        );


    bars

        .on(
            "mouseenter",
            function(event,d) {

                d3.select(this)
                    .attr(
                        "opacity",
                        0.7
                    );


                showTooltip(
                    event,
                    `
                    <strong>${d.country}</strong><br>
                    มูลค่า:
                    ${formatNumber(d.value)}
                    `
                );

            }
        )

        .on(
            "mousemove",
            moveTooltip
        )

        .on(
            "mouseleave",
            function() {

                d3.select(this)
                    .attr(
                        "opacity",
                        1
                    );


                hideTooltip();

            }
        )

        /*
            DRILLDOWN

            คลิกประเทศ
            แล้ว Filter Dashboard ทั้งหมด
        */

        .on(
            "click",
            function(event,d) {

                const countryFilter =
                    document.getElementById(
                        "countryFilter"
                    );


                countryFilter.value =
                    d.country;


                applyFilters();

            }
        );


    bars

        .transition()

        .duration(850)

        .ease(
            d3.easeCubicOut
        )

        .attr(
            "y",
            d => y(d.value)
        )

        .attr(
            "height",
            d =>
                height -
                margin.bottom -
                y(d.value)
        );


    /* X Axis */

    svg.append("g")

        .attr(
            "class",
            "axis"
        )

        .attr(
            "transform",
            `translate(
                0,
                ${height - margin.bottom}
            )`
        )

        .call(
            d3.axisBottom(x)
        )

        .selectAll("text")

        .attr(
            "transform",
            "rotate(-35)"
        )

        .style(
            "text-anchor",
            "end"
        );


    /* Y Axis */

    svg.append("g")

        .attr(
            "class",
            "axis"
        )

        .attr(
            "transform",
            `translate(
                ${margin.left},
                0
            )`
        )

        .call(
            d3.axisLeft(y)
                .ticks(5)
                .tickFormat(
                    formatAxisNumber
                )
        );

}


/* =====================================================
   DONUT CHART
===================================================== */

function drawDonutChart() {

    const container =
        document.getElementById(
            "donutChart"
        );


    container.innerHTML = "";


    if (
        filteredData.length === 0
    ) {

        showNoData(
            container
        );

        return;

    }


    const grouped =
        d3.rollups(

            filteredData,

            values =>
                values.length,

            d => d.tradeflow

        );


    const data =
        grouped.map(
            ([flow,value]) => ({
                flow,
                value
            })
        );


    const width =
        Math.max(
            container.clientWidth,
            320
        );


    const height =
        350;


    const radius =
        Math.min(
            width,
            height
        ) / 2 -
        35;


    const svg =
        d3.select(container)

            .append("svg")

            .attr(
                "viewBox",
                `0 0 ${width} ${height}`
            );


    const group =
        svg.append("g")

            .attr(
                "transform",
                `translate(
                    ${width / 2},
                    ${height / 2}
                )`
            );


    const pie =
        d3.pie()

            .sort(null)

            .value(
                d => d.value
            );


    const arc =
        d3.arc()

            .innerRadius(
                radius * 0.56
            )

            .outerRadius(
                radius
            );


    const arcHover =
        d3.arc()

            .innerRadius(
                radius * 0.56
            )

            .outerRadius(
                radius + 8
            );


    const paths =
        group.selectAll("path")

            .data(
                pie(data)
            )

            .enter()

            .append("path")

            .attr(
                "fill",
                (d,i) =>
                    pastelColors[
                        i %
                        pastelColors.length
                    ]
            )

            .attr(
                "stroke",
                "#ffffff"
            )

            .attr(
                "stroke-width",
                3
            );


    paths

        .on(
            "mouseenter",
            function(event,d) {

                d3.select(this)
                    .transition()
                    .duration(150)
                    .attr(
                        "d",
                        arcHover
                    );


                showTooltip(
                    event,
                    `
                    <strong>
                        Trade Flow ${d.data.flow}
                    </strong><br>
                    จำนวน:
                    ${formatNumber(
                        d.data.value
                    )}
                    รายการ
                    `
                );

            }
        )

        .on(
            "mousemove",
            moveTooltip
        )

        .on(
            "mouseleave",
            function() {

                d3.select(this)
                    .transition()
                    .duration(150)
                    .attr(
                        "d",
                        arc
                    );


                hideTooltip();

            }
        );


    paths

        .each(
            function(d) {

                this._current = {

                    startAngle: 0,

                    endAngle: 0

                };

            }
        )

        .transition()

        .duration(900)

        .attrTween(
            "d",
            function(d) {

                const interpolate =
                    d3.interpolate(
                        this._current,
                        d
                    );


                this._current =
                    interpolate(1);


                return function(t) {

                    return arc(
                        interpolate(t)
                    );

                };

            }
        );


    /* Center text */

    group.append("text")

        .attr(
            "text-anchor",
            "middle"
        )

        .attr(
            "dy",
            "-0.1em"
        )

        .style(
            "fill",
            "#60747d"
        )

        .style(
            "font-size",
            "24px"
        )

        .style(
            "font-weight",
            "600"
        )

        .text(
            formatNumber(
                filteredData.length
            )
        );


    group.append("text")

        .attr(
            "text-anchor",
            "middle"
        )

        .attr(
            "dy",
            "1.5em"
        )

        .style(
            "fill",
            "#9aa7ad"
        )

        .style(
            "font-size",
            "11px"
        )

        .text(
            "รายการทั้งหมด"
        );

}


/* =====================================================
   LINE CHART
===================================================== */

function drawLineChart() {

    const container =
        document.getElementById(
            "lineChart"
        );


    container.innerHTML = "";


    if (
        filteredData.length === 0
    ) {

        showNoData(
            container
        );

        return;

    }


    const grouped =
        d3.rollups(

            filteredData,

            values =>
                d3.sum(
                    values,
                    d => d.price
                ),

            d => d.month

        );


    const data =
        grouped

            .map(
                ([month,value]) => ({
                    month,
                    value
                })
            )

            .sort(
                (a,b) =>
                    a.month -
                    b.month
            );


    const width =
        Math.max(
            container.clientWidth,
            650
        );


    const height =
        420;


    const margin = {

        top: 25,

        right: 30,

        bottom: 55,

        left: 80

    };


    const svg =
        d3.select(container)

            .append("svg")

            .attr(
                "viewBox",
                `0 0 ${width} ${height}`
            );


    const x =
        d3.scaleLinear()

            .domain([
                1,
                12
            ])

            .range([
                margin.left,
                width - margin.right
            ]);


    const maxValue =
        d3.max(
            data,
            d => d.value
        ) || 1;


    const y =
        d3.scaleLinear()

            .domain([
                0,
                maxValue
            ])

            .nice()

            .range([
                height - margin.bottom,
                margin.top
            ]);


    const line =
        d3.line()

            .defined(
                d =>
                    Number.isFinite(
                        d.value
                    )
            )

            .x(
                d => x(d.month)
            )

            .y(
                d => y(d.value)
            )

            .curve(
                d3.curveMonotoneX
            );


    /* Grid */

    svg.append("g")

        .attr(
            "class",
            "grid"
        )

        .attr(
            "transform",
            `translate(${margin.left},0)`
        )

        .call(
            d3.axisLeft(y)
                .ticks(5)
                .tickSize(
                    -(width -
                        margin.left -
                        margin.right)
                )
                .tickFormat("")
        )

        .selectAll("line")

        .attr(
            "class",
            "grid-line"
        );


    /* Line */

    const path =
        svg.append("path")

            .datum(data)

            .attr(
                "fill",
                "none"
            )

            .attr(
                "stroke",
                "#74bef7"
            )

            .attr(
                "stroke-width",
                4
            )

            .attr(
                "stroke-linecap",
                "round"
            )

            .attr(
                "d",
                line
            );


    const pathLength =
        path.node()
            .getTotalLength();


    path

        .attr(
            "stroke-dasharray",
            `${pathLength} ${pathLength}`
        )

        .attr(
            "stroke-dashoffset",
            pathLength
        )

        .transition()

        .duration(1200)

        .ease(
            d3.easeCubicOut
        )

        .attr(
            "stroke-dashoffset",
            0
        );


    /* Points */

    const points =
        svg.selectAll(".line-point")

            .data(data)

            .enter()

            .append("circle")

            .attr(
                "class",
                "line-point"
            )

            .attr(
                "cx",
                d => x(d.month)
            )

            .attr(
                "cy",
                d => y(d.value)
            )

            .attr(
                "r",
                0
            )

            .attr(
                "fill",
                "#fd9097"
            )

            .attr(
                "stroke",
                "#ffffff"
            )

            .attr(
                "stroke-width",
                2
            );


    points

        .on(
            "mouseenter",
            function(event,d) {

                d3.select(this)
                    .transition()
                    .attr(
                        "r",
                        8
                    );


                showTooltip(
                    event,
                    `
                    <strong>
                        ${monthNames[d.month]}
                    </strong><br>
                    มูลค่า:
                    ${formatNumber(
                        d.value
                    )}
                    `
                );

            }
        )

        .on(
            "mousemove",
            moveTooltip
        )

        .on(
            "mouseleave",
            function() {

                d3.select(this)
                    .transition()
                    .attr(
                        "r",
                        5
                    );


                hideTooltip();

            }
        );


    points

        .transition()

        .delay(
            (_,i) => i * 60
        )

        .duration(400)

        .attr(
            "r",
            5
        );


    /* X Axis */

    svg.append("g")

        .attr(
            "class",
            "axis"
        )

        .attr(
            "transform",
            `translate(
                0,
                ${height - margin.bottom}
            )`
        )

        .call(
            d3.axisBottom(x)
                .ticks(12)
                .tickFormat(
                    d =>
                        monthNames[d]
                            ? monthNames[d]
                                .substring(0,3)
                            : d
                )
        );


    /* Y Axis */

    svg.append("g")

        .attr(
            "class",
            "axis"
        )

        .attr(
            "transform",
            `translate(
                ${margin.left},
                0
            )`
        )

        .call(
            d3.axisLeft(y)
                .ticks(5)
                .tickFormat(
                    formatAxisNumber
                )
        );


    /* =================================================
       ZOOM
    ================================================= */

    const zoom =
        d3.zoom()

            .scaleExtent([
                1,
                5
            ])

            .translateExtent([
                [0,0],
                [width,height]
            ])

            .extent([
                [0,0],
                [width,height]
            ])

            .on(
                "zoom",
                event => {

                    const newX =
                        event.transform
                            .rescaleX(x);


                    svg.select(
                        ".line-path"
                    );


                    path.attr(
                        "d",
                        line.x(
                            d =>
                                newX(
                                    d.month
                                )
                        )
                    );


                    points.attr(
                        "cx",
                        d =>
                            newX(
                                d.month
                            )
                    );


                    svg.select(
                        ".x-axis"
                    );

                }
            );


    svg.call(zoom);


    /* Double click reset */

    svg.on(
        "dblclick.zoom",
        null
    );


    svg.on(
        "dblclick",
        () => {

            svg
                .transition()
                .duration(500)
                .call(
                    zoom.transform,
                    d3.zoomIdentity
                );

        }
    );

}


/* =====================================================
   MAP INITIALIZATION
===================================================== */

function initializeMap() {

    const mapElement =
        document.getElementById(
            "map"
        );


    if (!mapElement) return;


    map =
        L.map(
            "map",
            {
                zoomControl: true
            }
        )
        .setView(
            [20,100],
            3
        );


    L.tileLayer(
        "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",
        {
            maxZoom: 18,

            attribution:
                "&copy; OpenStreetMap contributors"

        }
    ).addTo(map);

}


/* =====================================================
   UPDATE MAP
===================================================== */

function updateMap() {

    if (!map) return;


    mapMarkers.forEach(
        marker => {

            map.removeLayer(
                marker
            );

        }
    );


    mapMarkers = [];


    if (
        filteredData.length === 0
    ) {

        return;

    }


    const grouped =
        d3.rollups(

            filteredData,

            values =>
                d3.sum(
                    values,
                    d => d.price
                ),

            d => d.country

        );


    const data =
        grouped

            .map(
                ([country,value]) => ({
                    country,
                    value
                })
            )

            .sort(
                (a,b) =>
                    b.value -
                    a.value
            )

            .slice(
                0,
                30
            );


    const maxValue =
        d3.max(
            data,
            d => d.value
        ) || 1;


    data.forEach(
        d => {

            const coordinates =
                countryCoordinates[
                    d.country
                ];


            if (!coordinates) {

                return;

            }


            const radius =
                7 +
                (
                    Math.sqrt(
                        d.value /
                        maxValue
                    ) * 24
                );


            const marker =
                L.circleMarker(
                    coordinates,
                    {

                        radius,

                        fillColor:
                            "#fd9097",

                        color:
                            "#ffffff",

                        weight:
                            2,

                        opacity:
                            1,

                        fillOpacity:
                            0.75

                    }
                );


            marker.bindTooltip(

                `
                <div class="map-popup">

                    <strong>
                        ${d.country}
                    </strong>

                    <br>

                    มูลค่าการค้า:
                    ${formatNumber(
                        d.value
                    )}

                </div>
                `,

                {
                    direction:
                        "top",

                    sticky:
                        true

                }

            );


            marker.addTo(map);


            mapMarkers.push(
                marker
            );

        }
    );

}


/* =====================================================
   TABLE
===================================================== */

function getSearchData() {

    const input =
        document.getElementById(
            "searchInput"
        );


    const keyword =
        input.value
            .trim()
            .toLowerCase();


    if (!keyword) {

        return filteredData;

    }


    return filteredData.filter(
        d => {

            return (

                d.country
                    .toLowerCase()
                    .includes(
                        keyword
                    )

                ||

                d.fish
                    .toLowerCase()
                    .includes(
                        keyword
                    )

                ||

                d.productTH
                    .toLowerCase()
                    .includes(
                        keyword
                    )

            );

        }
    );

}


function updateTable() {

    const tbody =
        document.getElementById(
            "dataTable"
        );


    tbody.innerHTML = "";


    const searchData =
        getSearchData();


    const totalPages =
        Math.max(
            1,
            Math.ceil(
                searchData.length /
                ROWS_PER_PAGE
            )
        );


    if (
        currentPage >
        totalPages
    ) {

        currentPage =
            totalPages;

    }


    const start =
        (
            currentPage -
            1
        ) *
        ROWS_PER_PAGE;


    const pageData =
        searchData.slice(
            start,
            start +
            ROWS_PER_PAGE
        );


    pageData.forEach(
        d => {

            const row =
                document.createElement(
                    "tr"
                );


            row.innerHTML = `

                <td>
                    ${d.year || "-"}
                </td>

                <td>
                    ${
                        monthNames[d.month]
                        || d.month
                        || "-"
                    }
                </td>

                <td>
                    ${escapeHTML(
                        d.country
                    )}
                </td>

                <td>
                    ${formatNumber(
                        d.weight
                    )}
                </td>

                <td>
                    ${formatNumber(
                        d.quantity
                    )}
                </td>

                <td>
                    ${formatNumber(
                        d.price
                    )}
                </td>

                <td>
                    ${escapeHTML(
                        d.fish
                    )}
                </td>

            `;


            tbody.appendChild(
                row
            );

        }
    );


    document.getElementById(
        "pageInfo"
    ).textContent =
        `หน้า ${currentPage} / ${totalPages}`;


    document.getElementById(
        "prevPage"
    ).disabled =
        currentPage <= 1;


    document.getElementById(
        "nextPage"
    ).disabled =
        currentPage >= totalPages;

}


/* =====================================================
   TOOLTIP FUNCTIONS
===================================================== */

function showTooltip(
    event,
    html
) {

    tooltip

        .html(html)

        .style(
            "opacity",
            1
        );


    moveTooltip(event);

}


function moveTooltip(event) {

    tooltip

        .style(
            "left",
            `${event.clientX + 14}px`
        )

        .style(
            "top",
            `${event.clientY + 14}px`
        );

}


function hideTooltip() {

    tooltip

        .style(
            "opacity",
            0
        );

}


/* =====================================================
   NO DATA
===================================================== */

function showNoData(
    container
) {

    container.innerHTML = `

        <div style="
            height:100%;
            display:flex;
            align-items:center;
            justify-content:center;
            color:#9aa7ad;
            font-size:14px;
        ">

            📭 ไม่พบข้อมูลตามตัวกรองที่เลือก

        </div>

    `;

}


/* =====================================================
   FORMAT NUMBER
===================================================== */

function formatNumber(
    value
) {

    if (
        !Number.isFinite(
            Number(value)
        )
    ) {

        return "0";

    }


    return Number(value)
        .toLocaleString(
            "th-TH",
            {
                maximumFractionDigits: 2
            }
        );

}


function formatAxisNumber(
    value
) {

    if (
        Math.abs(value) >=
        1000000000
    ) {

        return (
            value / 1000000000
        ).toFixed(1) + "B";

    }


    if (
        Math.abs(value) >=
        1000000
    ) {

        return (
            value / 1000000
        ).toFixed(1) + "M";

    }


    if (
        Math.abs(value) >=
        1000
    ) {

        return (
            value / 1000
        ).toFixed(1) + "K";

    }


    return value;

}


/* =====================================================
   ESCAPE HTML
===================================================== */

function escapeHTML(
    value
) {

    return String(value)
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