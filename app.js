/* =========================================================
   FISH TRADE INTERACTIVE DASHBOARD
   JavaScript Framework : Vue 3
   Excel Reader         : SheetJS
   Map                  : Leaflet
   D3.js                : NOT USED
========================================================= */

const { createApp } = Vue;


/* =========================================================
   VUE APPLICATION
========================================================= */

createApp({

    /* =====================================================
       DATA
    ===================================================== */

    data() {

        return {

            /* -------------------------
               Raw data
            -------------------------- */

            rawData: [],


            /* -------------------------
               Loading
            -------------------------- */

            loading: true,

            errorMessage: "",


            /* -------------------------
               Filters
            -------------------------- */

            filters: {

                year: "",

                month: "",

                country: "",

                tradeflow: ""

            },


            /* -------------------------
               Search
            -------------------------- */

            searchText: "",


            /* -------------------------
               Pagination
            -------------------------- */

            currentPage: 1,

            pageSize: 10,


            /* -------------------------
               Map
            -------------------------- */

            map: null,

            mapMarkers: [],


            /* -------------------------
               Chart colors
            -------------------------- */

            chartColors: [

                "#fd9097",

                "#f8c891",

                "#f8f862",

                "#6ef88c",

                "#74bef7",

                "#b7a7ff",

                "#ffb6d9",

                "#8dd3c7",

                "#ffd166",

                "#9ad0f5"

            ],


            /* -------------------------
               Country coordinates
            -------------------------- */

            countryCoordinates: {

                "ไทย": [15.8700, 100.9925],

                "เวียดนาม": [14.0583, 108.2772],

                "ญี่ปุ่น": [36.2048, 138.2529],

                "จีน": [35.8617, 104.1954],

                "เกาหลีใต้": [35.9078, 127.7669],

                "เกาหลีเหนือ": [40.3399, 127.5101],

                "สหรัฐอเมริกา": [37.0902, -95.7129],

                "สหรัฐฯ": [37.0902, -95.7129],

                "ออสเตรเลีย": [-25.2744, 133.7751],

                "นิวซีแลนด์": [-40.9006, 174.8860],

                "สิงคโปร์": [1.3521, 103.8198],

                "มาเลเซีย": [4.2105, 101.9758],

                "อินโดนีเซีย": [-0.7893, 113.9213],

                "ฟิลิปปินส์": [12.8797, 121.7740],

                "อินเดีย": [20.5937, 78.9629],

                "มัลดีฟส์": [3.2028, 73.2207],

                "ฝรั่งเศส": [46.2276, 2.2137],

                "เยอรมนี": [51.1657, 10.4515],

                "อังกฤษ": [55.3781, -3.4360],

                "สหราชอาณาจักร": [55.3781, -3.4360],

                "แคนาดา": [56.1304, -106.3468],

                "รัสเซีย": [61.5240, 105.3188],

                "บราซิล": [-14.2350, -51.9253],

                "แอฟริกาใต้": [-30.5595, 22.9375],

                "สเปน": [40.4637, -3.7492],

                "อิตาลี": [41.8719, 12.5674],

                "เนเธอร์แลนด์": [52.1326, 5.2913],

                "นอร์เวย์": [60.4720, 8.4689],

                "สวีเดน": [60.1282, 18.6435],

                "เดนมาร์ก": [56.2639, 9.5018],

                "เม็กซิโก": [23.6345, -102.5528],

                "ชิลี": [-35.6751, -71.5430]

            }

        };

    },


    /* =====================================================
       COMPUTED
    ===================================================== */

    computed: {

        /* =================================================
           YEARS
        ================================================= */

        years() {

            return [

                ...new Set(

                    this.rawData

                        .map(row => row.year)

                        .filter(value => value !== "")

                )

            ].sort((a, b) => a - b);

        },


        /* =================================================
           MONTHS
        ================================================= */

        months() {

            const monthOrder = [

                "January",
                "February",
                "March",
                "April",
                "May",
                "June",
                "July",
                "August",
                "September",
                "October",
                "November",
                "December"

            ];


            const existing = [

                ...new Set(

                    this.rawData

                        .map(row => row.month)

                        .filter(value => value !== "")

                )

            ];


            return existing.sort((a, b) => {

                const ai = monthOrder.indexOf(
                    String(a)
                );

                const bi = monthOrder.indexOf(
                    String(b)
                );

                if (ai !== -1 && bi !== -1) {

                    return ai - bi;

                }

                return String(a).localeCompare(
                    String(b)
                );

            });

        },


        /* =================================================
           COUNTRIES
        ================================================= */

        countries() {

            return [

                ...new Set(

                    this.rawData

                        .map(row => row.country)

                        .filter(value => value !== "")

                )

            ].sort(

                (a, b) => String(a).localeCompare(
                    String(b),
                    "th"
                )

            );

        },


        /* =================================================
           TRADE FLOWS
        ================================================= */

        tradeflows() {

            return [

                ...new Set(

                    this.rawData

                        .map(row => row.tradeflow)

                        .filter(value => value !== "")

                )

            ];

        },


        /* =================================================
           FILTERED DATA
        ================================================= */

        filteredData() {

            let data = this.rawData;


            /* Year */

            if (this.filters.year !== "") {

                data = data.filter(

                    row => String(row.year) ===
                           String(this.filters.year)

                );

            }


            /* Month */

            if (this.filters.month !== "") {

                data = data.filter(

                    row => String(row.month) ===
                           String(this.filters.month)

                );

            }


            /* Country */

            if (this.filters.country !== "") {

                data = data.filter(

                    row => String(row.country) ===
                           String(this.filters.country)

                );

            }


            /* Trade Flow */

            if (this.filters.tradeflow !== "") {

                data = data.filter(

                    row => String(row.tradeflow) ===
                           String(this.filters.tradeflow)

                );

            }


            return data;

        },


        /* =================================================
           SEARCHED DATA
        ================================================= */

        searchedData() {

            const keyword = this.searchText
                .trim()
                .toLowerCase();


            if (!keyword) {

                return this.filteredData;

            }


            return this.filteredData.filter(row => {

                const country =
                    String(row.country || "")
                        .toLowerCase();

                const fish =
                    String(row.fish || "")
                        .toLowerCase();

                const productTH =
                    String(row.productTH || "")
                        .toLowerCase();

                const productEN =
                    String(row.productEN || "")
                        .toLowerCase();


                return (

                    country.includes(keyword) ||

                    fish.includes(keyword) ||

                    productTH.includes(keyword) ||

                    productEN.includes(keyword)

                );

            });

        },


        /* =================================================
           KPI
        ================================================= */

        kpi() {

            return {

                records: this.filteredData.length,

                weight: this.sum(
                    this.filteredData,
                    "weight"
                ),

                quantity: this.sum(
                    this.filteredData,
                    "quantity"
                ),

                price: this.sum(
                    this.filteredData,
                    "price"
                )

            };

        },


        /* =================================================
           BAR DATA
        ================================================= */

        barData() {

            const map = new Map();


            this.filteredData.forEach(row => {

                const country =
                    row.country || "ไม่ระบุ";

                const value =
                    Number(row.price) || 0;


                map.set(

                    country,

                    (map.get(country) || 0) + value

                );

            });


            return Array.from(map.entries())

                .map(([country, value]) => ({

                    country,

                    value

                }))

                .sort(

                    (a, b) => b.value - a.value

                )

                .slice(0, 10);

        },


        /* =================================================
           DONUT DATA
        ================================================= */

        donutData() {

            const map = new Map();


            this.filteredData.forEach(row => {

                const label =
                    this.tradeflowText(
                        row.tradeflow
                    );

                map.set(

                    label,

                    (map.get(label) || 0) + 1

                );

            });


            const total =
                this.filteredData.length;


            return Array.from(map.entries())

                .map(([label, value]) => ({

                    label,

                    value,

                    percent:
                        total === 0
                            ? 0
                            : (value / total) * 100

                }))

                .sort(

                    (a, b) => b.value - a.value

                );

        },


        /* =================================================
           DONUT STYLE
        ================================================= */

        donutStyle() {

            if (this.donutData.length === 0) {

                return {

                    background:
                        "#eeeeee"

                };

            }


            let current = 0;


            const parts = [];


            this.donutData.forEach(

                (item, index) => {

                    const start = current;

                    current += item.percent;

                    const end = current;


                    parts.push(

                        `${this.chartColors[
                            index %
                            this.chartColors.length
                        ]} ${start}% ${end}%`

                    );

                }

            );


            return {

                background:
                    `conic-gradient(${parts.join(", ")})`

            };

        },


        /* =================================================
           LINE DATA
        ================================================= */

        lineData() {

            const monthMap = new Map();


            this.filteredData.forEach(row => {

                const month =
                    this.monthNumber(
                        row.month
                    );


                const value =
                    Number(row.price) || 0;


                if (!monthMap.has(month)) {

                    monthMap.set(month, {

                        month,

                        value: 0

                    });

                }


                monthMap.get(month).value += value;

            });


            return Array.from(
                monthMap.values()
            )

                .sort(
                    (a, b) => a.month - b.month
                );

        },


        /* =================================================
           LINE POINTS
        ================================================= */

        linePoints() {

            const data = this.lineData;


            if (data.length === 0) {

                return [];

            }


            const width = 790;

            const left = 70;

            const right = 860;

            const top = 30;

            const bottom = 290;


            const maxValue =
                Math.max(
                    ...data.map(d => d.value),
                    1
                );


            return data.map((item, index) => {

                let x;


                if (data.length === 1) {

                    x = (left + right) / 2;

                } else {

                    x =
                        left +
                        (
                            index /
                            (data.length - 1)
                        ) *
                        (right - left);

                }


                const y =
                    bottom -
                    (
                        item.value /
                        maxValue
                    ) *
                    (bottom - top);


                return {

                    month:
                        this.monthName(
                            item.month
                        ),

                    value:
                        item.value,

                    x,

                    y

                };

            });

        },


        /* =================================================
           LINE POLYLINE
        ================================================= */

        linePolyline() {

            return this.linePoints

                .map(
                    point =>
                        `${point.x},${point.y}`
                )

                .join(" ");

        },


        /* =================================================
           LINE AREA
        ================================================= */

        lineAreaPoints() {

            if (this.linePoints.length === 0) {

                return "";

            }


            const first =
                this.linePoints[0];

            const last =
                this.linePoints[
                    this.linePoints.length - 1
                ];


            const line =
                this.linePoints

                    .map(
                        point =>
                            `${point.x},${point.y}`
                    )

                    .join(" ");


            return `

                ${first.x},290

                ${line}

                ${last.x},290

            `;

        },


        /* =================================================
           LINE GRID
        ================================================= */

        lineGrid() {

            return [

                30,
                95,
                160,
                225,
                290

            ];

        },


        /* =================================================
           LINE Y LABELS
        ================================================= */

        lineYLabels() {

            const maxValue =
                Math.max(

                    ...this.lineData.map(
                        d => d.value
                    ),

                    1

                );


            return [

                {
                    value: maxValue,
                    y: 30
                },

                {
                    value: maxValue * 0.75,
                    y: 95
                },

                {
                    value: maxValue * 0.50,
                    y: 160
                },

                {
                    value: maxValue * 0.25,
                    y: 225
                },

                {
                    value: 0,
                    y: 290
                }

            ];

        },


        /* =================================================
           TOTAL PAGES
        ================================================= */

        totalPages() {

            return Math.max(

                1,

                Math.ceil(

                    this.searchedData.length /
                    this.pageSize

                )

            );

        },


        /* =================================================
           PAGINATED ROWS
        ================================================= */

        paginatedRows() {

            const start =
                (
                    this.currentPage - 1
                ) *
                this.pageSize;


            return this.searchedData.slice(

                start,

                start + this.pageSize

            );

        }

    },


    /* =====================================================
       WATCH
    ===================================================== */

    watch: {

        filters: {

            deep: true,

            handler() {

                this.currentPage = 1;

                this.$nextTick(() => {

                    this.updateMap();

                });

            }

        },


        searchText() {

            this.currentPage = 1;

        }

    },


    /* =====================================================
       MOUNTED
    ===================================================== */

    mounted() {

        this.loadExcel();

    },


    /* =====================================================
       METHODS
    ===================================================== */

    methods: {


        /* =================================================
           LOAD EXCEL
        ================================================= */

        async loadExcel() {

            this.loading = true;

            this.errorMessage = "";


            try {

                const response =
                    await fetch("data.xlsx");


                if (!response.ok) {

                    throw new Error(

                        "ไม่สามารถเปิดไฟล์ data.xlsx ได้"

                    );

                }


                const buffer =
                    await response.arrayBuffer();


                const workbook =
                    XLSX.read(

                        buffer,

                        {
                            type: "array"
                        }

                    );


                const sheetName =
                    workbook.SheetNames.includes(
                        "Cleaned_Data"
                    )
                        ? "Cleaned_Data"
                        : workbook.SheetNames[0];


                const worksheet =
                    workbook.Sheets[sheetName];


                const rows =
                    XLSX.utils.sheet_to_json(

                        worksheet,

                        {
                            defval: ""
                        }

                    );


                this.rawData =
                    rows.map(

                        (row, index) =>
                            this.normalizeRow(
                                row,
                                index
                            )

                    );


                this.loading = false;


                this.$nextTick(() => {

                    this.initMap();

                    this.updateMap();

                });


            } catch (error) {

                console.error(error);

                this.loading = false;

                this.errorMessage =
                    error.message ||
                    "เกิดข้อผิดพลาดในการอ่านข้อมูล";

            }

        },


        /* =================================================
           NORMALIZE ROW
        ================================================= */

        normalizeRow(row, index) {

            return {

                _id: index + 1,


                year:
                    this.cleanValue(
                        row.year
                    ),


                month:
                    this.cleanValue(
                        row.month
                    ),


                heading11:
                    this.cleanValue(
                        row.heading11
                    ),


                countryID:
                    this.cleanValue(
                        row.countryID
                    ),


                country:
                    this.cleanValue(

                        row.countryNameTH ??
                        row.country ??
                        row.Country

                    ),


                weight:
                    this.toNumber(
                        row.weight
                    ),


                quantity:
                    this.toNumber(
                        row.quantity
                    ),


                price:
                    this.toNumber(
                        row.price
                    ),


                tradeflow:
                    this.cleanValue(
                        row.tradeflow
                    ),


                productEN:
                    this.cleanValue(

                        row.productDetailEN ??
                        row.productEN

                    ),


                fish:
                    this.cleanValue(

                        row.fishName ??
                        row.fish

                    ),


                productTH:
                    this.cleanValue(

                        row.productDetailTH ??
                        row.productTH

                    )

            };

        },


        /* =================================================
           CLEAN VALUE
        ================================================= */

        cleanValue(value) {

            if (
                value === null ||
                value === undefined
            ) {

                return "";

            }


            return String(value).trim();

        },


        /* =================================================
           TO NUMBER
        ================================================= */

        toNumber(value) {

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

        },


        /* =================================================
           SUM
        ================================================= */

        sum(data, field) {

            return data.reduce(

                (total, row) => {

                    return total +
                        (
                            Number(row[field]) || 0
                        );

                },

                0

            );

        },


        /* =================================================
           FORMAT NUMBER
        ================================================= */

        formatNumber(value) {

            return new Intl.NumberFormat(
                "th-TH",
                {
                    maximumFractionDigits: 2
                }
            ).format(
                Number(value) || 0
            );

        },


        /* =================================================
           FORMAT COMPACT
        ================================================= */

        formatCompact(value) {

            const number =
                Number(value) || 0;


            if (number >= 1000000000) {

                return (
                    number / 1000000000
                ).toFixed(1) + "B";

            }


            if (number >= 1000000) {

                return (
                    number / 1000000
                ).toFixed(1) + "M";

            }


            if (number >= 1000) {

                return (
                    number / 1000
                ).toFixed(1) + "K";

            }


            return number.toFixed(0);

        },


        /* =================================================
           RESET
        ================================================= */

        resetFilters() {

            this.filters = {

                year: "",

                month: "",

                country: "",

                tradeflow: ""

            };


            this.searchText = "";

            this.currentPage = 1;

        },


        /* =================================================
           SELECT COUNTRY
        ================================================= */

        selectCountry(country) {

            this.filters.country = country;

            this.currentPage = 1;

        },


        /* =================================================
           BAR WIDTH
        ================================================= */

        barWidth(value) {

            if (
                this.barData.length === 0
            ) {

                return "0%";

            }


            const max =
                Math.max(

                    ...this.barData.map(
                        d => d.value
                    ),

                    1

                );


            return (
                Math.max(
                    3,
                    (value / max) * 100
                )
            ) + "%";

        },


        /* =================================================
           BAR COLOR
        ================================================= */

        barColor(index) {

            return this.chartColors[
                index %
                this.chartColors.length
            ];

        },


        /* =================================================
           MONTH NUMBER
        ================================================= */

        monthNumber(value) {

            const text =
                String(value || "")
                    .trim()
                    .toLowerCase();


            const monthMap = {

                january: 1,
                jan: 1,

                february: 2,
                feb: 2,

                march: 3,
                mar: 3,

                april: 4,
                apr: 4,

                may: 5,

                june: 6,
                jun: 6,

                july: 7,
                jul: 7,

                august: 8,
                aug: 8,

                september: 9,
                sep: 9,

                october: 10,
                oct: 10,

                november: 11,
                nov: 11,

                december: 12,
                dec: 12

            };


            if (monthMap[text]) {

                return monthMap[text];

            }


            const number =
                Number(value);


            if (
                Number.isFinite(number) &&
                number >= 1 &&
                number <= 12
            ) {

                return number;

            }


            return 0;

        },


        /* =================================================
           MONTH NAME
        ================================================= */

        monthName(number) {

            const names = [

                "Jan",
                "Feb",
                "Mar",
                "Apr",
                "May",
                "Jun",
                "Jul",
                "Aug",
                "Sep",
                "Oct",
                "Nov",
                "Dec"

            ];


            return names[number - 1] || String(number);

        },


        /* =================================================
           TRADE FLOW TEXT
        ================================================= */

        tradeflowText(value) {

            const text =
                String(value || "")
                    .trim();


            if (text === "1") {

                return "นำเข้า";

            }


            if (text === "2") {

                return "ส่งออก";

            }


            if (
                text.toLowerCase() ===
                "import"
            ) {

                return "นำเข้า";

            }


            if (
                text.toLowerCase() ===
                "export"
            ) {

                return "ส่งออก";

            }


            return text || "ไม่ระบุ";

        },


        /* =================================================
           TRADE CLASS
        ================================================= */

        tradeClass(value) {

            const text =
                String(value || "")
                    .toLowerCase();


            if (
                text === "1" ||
                text.includes("import") ||
                text.includes("นำเข้า")
            ) {

                return "import";

            }


            if (
                text === "2" ||
                text.includes("export") ||
                text.includes("ส่งออก")
            ) {

                return "export";

            }


            return "other";

        },


        /* =================================================
           PREVIOUS PAGE
        ================================================= */

        previousPage() {

            if (
                this.currentPage > 1
            ) {

                this.currentPage--;

            }

        },


        /* =================================================
           NEXT PAGE
        ================================================= */

        nextPage() {

            if (
                this.currentPage <
                this.totalPages
            ) {

                this.currentPage++;

            }

        },


        /* =================================================
           INIT MAP
        ================================================= */

        initMap() {

            if (
                this.map ||
                !document.getElementById("map")
            ) {

                return;

            }


            this.map =
                L.map("map");


            this.map.setView(

                [15.8700, 100.9925],

                3

            );


            L.tileLayer(

                "https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",

                {

                    attribution:
                        "&copy; OpenStreetMap contributors",

                    maxZoom: 18

                }

            ).addTo(this.map);

        },


        /* =================================================
           UPDATE MAP
        ================================================= */

        updateMap() {

            if (!this.map) {

                return;

            }


            /* Remove old markers */

            this.mapMarkers.forEach(

                marker => {

                    this.map.removeLayer(
                        marker
                    );

                }

            );


            this.mapMarkers = [];


            /* Group data */

            const countryMap = new Map();


            this.filteredData.forEach(row => {

                const country =
                    row.country;


                if (!country) {

                    return;

                }


                if (!countryMap.has(country)) {

                    countryMap.set(

                        country,

                        {

                            country,

                            value: 0,

                            quantity: 0,

                            weight: 0

                        }

                    );

                }


                const item =
                    countryMap.get(country);


                item.value +=
                    Number(row.price) || 0;

                item.quantity +=
                    Number(row.quantity) || 0;

                item.weight +=
                    Number(row.weight) || 0;

            });


            /* Create markers */

            countryMap.forEach(item => {

                const coordinates =
                    this.findCoordinates(
                        item.country
                    );


                if (!coordinates) {

                    return;

                }


                const radius =
                    Math.max(

                        5,

                        Math.min(

                            35,

                            Math.sqrt(
                                item.value
                            ) / 10

                        )

                    );


                const marker =
                    L.circleMarker(

                        coordinates,

                        {

                            radius,

                            weight: 2,

                            color: "#ffffff",

                            fillColor: "#74bef7",

                            fillOpacity: 0.75

                        }

                    );


                marker.bindPopup(`

                    <div class="map-popup">

                        <h3>
                            ${this.escapeHTML(
                                item.country
                            )}
                        </h3>

                        <p>
                            <strong>
                                มูลค่าการค้า:
                            </strong>
                            ${this.formatNumber(
                                item.value
                            )}
                        </p>

                        <p>
                            <strong>
                                น้ำหนัก:
                            </strong>
                            ${this.formatNumber(
                                item.weight
                            )}
                        </p>

                        <p>
                            <strong>
                                ปริมาณ:
                            </strong>
                            ${this.formatNumber(
                                item.quantity
                            )}
                        </p>

                    </div>

                `);


                marker.addTo(this.map);


                this.mapMarkers.push(
                    marker
                );

            });

        },


        /* =================================================
           FIND COORDINATES
        ================================================= */

        findCoordinates(country) {

            if (
                this.countryCoordinates[
                    country
                ]
            ) {

                return this.countryCoordinates[
                    country
                ];

            }


            const normalized =
                String(country || "")
                    .trim();


            const key =
                Object.keys(
                    this.countryCoordinates
                ).find(

                    name =>
                        name.includes(normalized) ||
                        normalized.includes(name)

                );


            return key
                ? this.countryCoordinates[key]
                : null;

        },


        /* =================================================
           ESCAPE HTML
        ================================================= */

        escapeHTML(value) {

            return String(value || "")

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

    }

}).mount("#app");