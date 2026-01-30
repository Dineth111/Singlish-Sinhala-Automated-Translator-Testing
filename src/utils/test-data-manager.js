const XLSX = require('xlsx');
const { utils, writeFile } = XLSX;
const { dirname } = require('path');
const { existsSync, mkdirSync } = require('fs-extra');
const fs = require('fs');
const path = require('path');

class TestCaseManager {
    constructor() {
        this.dataList = [];
    }

    // Gather everything into one list
    collectAllData() {
        this.dataList = [];

        this.setupPositiveData();
        this.setupNegativeData();
        this.setupUIData();

        console.log(`Gathered ${this.dataList.length} total cases`);
        return this.dataList;
    }

    setupPositiveData() {
        console.log("Setting up positive functional data...");

        const positiveCases = [
            {
                id: "P_TC_01",
                name: "Simple sentence test",
                length: "S",
                input: "mama gedhara yanavaa.",
                expected: "මම ගෙදර යනවා.",
                justification: "Basic sentence with present tense converts correctly. Spacing and word order are accurate.",
                category: "Daily language usage\nSimple sentence\nS\nAccuracy validation"
            },
            {
                id: "P_TC_02",
                name: "Compound sentence test",
                length: "M",
                input: "mama gedhara yanavaa, haebaevi vahina nisaa dhaenna yannee naee.",
                expected: "මම ගෙදර යනවා, හැබැවි වහින නිසා දැන්න යන්නේ නෑ.",
                justification: "Testing a longer sentence with two parts joined by a comma. The system handles the conjunction correctly.",
                category: "Daily language usage\nCompound sentence\nM\nAccuracy validation"
            },
            {
                id: "P_TC_03",
                name: "Conditional sentence test",
                length: "M",
                input: "oya enavaanam mama balan innavaa.",
                expected: "ඔය එනවානම් මම බලන් ඉන්නවා.",
                justification: "Checking if the conditional form (enavaanam) is parsed as a single unit or split up. Works fine here.",
                category: "Daily language usage\nComplex sentence\nM\nAccuracy validation"
            },
            {
                id: "P_TC_04",
                name: "Interrogative sentence conversion",
                length: "S",
                input: "oyaata kohomadha?",
                expected: "ඔයාට කොහොමද?",
                justification: "Question form correctly converted with question mark preserved.",
                category: "Greeting / request / response\nInterrogative (question)\nS\nAccuracy validation"
            },
            {
                id: "P_TC_05",
                name: "Command sentence conversion",
                length: "S",
                input: "vahaama enna.",
                expected: "වහාම එන්න.",
                justification: "Command form accurately converted to Sinhala imperative.",
                category: "Daily language usage\nImperative (command)\nS\nAccuracy validation"
            },
            {
                id: "P_TC_06",
                name: "Future tense test",
                length: "S",
                input: "api heta enavaa.",
                expected: "අපි හෙට එනවා.",
                justification: "Future tense correctly converted with proper time reference.",
                category: "Daily language usage\nFuture tense\nS\nAccuracy validation"
            },
            {
                id: "P_TC_07",
                name: "Negative sentence test",
                length: "S",
                input: "api heta ennee naehae",
                expected: "අපි හෙට එන්නේ නැහැ",
                justification: "Negative sentence correctly converted with negation marker.",
                category: "Daily language usage\nNegation (negative form)\nS\nAccuracy validation"
            },
            {
                id: "P_TC_08",
                name: "Greeting conversion",
                length: "S",
                input: "aayuboovan!",
                expected: "ආයුබෝවන්!",
                justification: "Standard greeting accurately translated.",
                category: "Greeting / request / response\nSimple sentence\nS\nAccuracy validation"
            },
            {
                id: "P_TC_09",
                name: "Polite request test",
                length: "M",
                input: "karuNaakaralaa mata podi udhavvak karanna puLuvandha?",
                expected: "කරුණාකරලා මට පොඩි උදව්වක් කරන්න පුළුවන්ද?",
                justification: "Polite request form correctly handled by the translator.",
                category: "Greeting / request / response\nInterrogative (question)\nM\nAccuracy validation"
            },
            {
                id: "P_TC_10",
                name: "Informal phrasing test",
                length: "S",
                input: "eeyi, ooka dhiyan.",
                expected: "ඒයි, ඕක දියන්.",
                justification: "Informal colloquial terms correctly converted.",
                category: "Slang / informal language\nSimple sentence\nS\nAccuracy validation"
            },
            {
                id: "P_TC_11",
                name: "Daily expression test",
                length: "S",
                input: "mata nidhimathayi.",
                expected: "මට නිදිමතයි.",
                justification: "Common daily expression accurately converted.",
                category: "Daily language usage\nSimple sentence\nS\nAccuracy validation"
            },
            {
                id: "P_TC_12",
                name: "Multi-word text conversion",
                length: "M",
                input: "mata oona poddak inna hariyata vaeda gihin enna kaeema kanna baya naee",
                expected: "මට ඕන පොඩ්ඩක් ඉන්න හරියට වැඩ ගිහින් එන්න කෑම කන්න බය නෑ",
                justification: "Multi-word collocation correctly converted.",
                category: "Word combination / phrase pattern\nComplex sentence\nM\nAccuracy validation"
            },
            {
                id: "P_TC_13",
                name: "No-space input test",
                length: "S",
                input: "mamagedharayanavaa",
                expected: "මමගෙදරයනවා",
                justification: "Joined words without spaces correctly interpreted.",
                category: "Formatting (spaces / line breaks / paragraph)\nSimple sentence\nS\nRobustness validation"
            },
            {
                id: "P_TC_14",
                name: "Repeated words test",
                length: "S",
                input: "hari hari eka eka",
                expected: "හරි හරි එක එක",
                justification: "Repeated words for emphasis correctly converted.",
                category: "Word combination / phrase pattern\nSimple sentence\nS\nAccuracy validation"
            },
            {
                id: "P_TC_15",
                name: "Past tense test",
                length: "S",
                input: "mama iyee gedhara giyaa.",
                expected: "මම ඉයේ ගෙදර ගියා.",
                justification: "Past tense correctly converted with proper verb conjugation.",
                category: "Daily language usage\nPast tense\nS\nAccuracy validation"
            },
            {
                id: "P_TC_16",
                name: "Informal negation test",
                length: "S",
                input: "mata eeka karanna baee.",
                expected: "මට ඒක කරන්න බෑ.",
                justification: "Alternative negation form 'baee' correctly handled.",
                category: "Daily language usage\nNegation (negative form)\nS\nAccuracy validation"
            },
            {
                id: "P_TC_17",
                name: "Singular pronoun test",
                length: "S",
                input: "eyaa gedhara giyaa.",
                expected: "එයා ගෙදර ගියා.",
                justification: "Singular third-person pronoun correctly converted.",
                category: "Daily language usage\nPronoun variation (I/you/we/they)\nS\nAccuracy validation"
            },
            {
                id: "P_TC_18",
                name: "Plural pronoun test",
                length: "S",
                input: "eyaalaa enavaa.",
                expected: "එයාලා එනවා.",
                justification: "Plural marker 'laa' correctly converted to Sinhala plural form.",
                category: "Daily language usage\nPlural form\nS\nAccuracy validation"
            },
            {
                id: "P_TC_19",
                name: "Imperative request test",
                length: "S",
                input: "eeka dhenna.",
                expected: "ඒක දෙන්න.",
                justification: "Direct imperative request correctly converted.",
                category: "Greeting / request / response\nImperative (command)\nS\nAccuracy validation"
            },
            {
                id: "P_TC_20",
                name: "Mixed language technical test",
                length: "M",
                input: "Zoom meeting ekak thiyennee. mama link eka WhatsApp karanna oone.",
                expected: "Zoom meeting එකක් තියෙන්නේ. මම link එක WhatsApp කරන්න ඕනෙ.",
                justification: "Testing mixed input where English technical words stay as they are.",
                category: "Mixed Singlish + English\nCompound sentence\nM\nRobustness validation"
            },
            {
                id: "P_TC_21",
                name: "English place name test",
                length: "S",
                input: "api trip eka Kandy valata yamudha.",
                expected: "අපි trip එක Kandy වලට යමුද.",
                justification: "English proper noun 'Kandy' correctly preserved.",
                category: "Names / places / common English words\nInterrogative (question)\nS\nAccuracy validation"
            },
            {
                id: "P_TC_22",
                name: "English abbreviation test",
                length: "S",
                input: "mata OTP eka yanna oone.",
                expected: "මට OTP එක යන්න ඕනෙ.",
                justification: "English abbreviation 'OTP' correctly preserved in context.",
                category: "Mixed Singlish + English\nSimple sentence\nS\nRobustness validation"
            },
            {
                id: "P_TC_23",
                name: "Punctuation validation",
                length: "S",
                input: "oyaath enavadha? (hithana)",
                expected: "ඔයාත් එනවද? (හිතන)",
                justification: "Parentheses and question mark correctly preserved.",
                category: "Punctuation / numbers\nInterrogative (question)\nS\nAccuracy validation"
            },
            {
                id: "P_TC_24",
                name: "Date and Time format test",
                length: "S",
                input: "dhesaembar 25 7.30 AM",
                expected: "දෙසැම්බර් 25 7.30 AM",
                justification: "Date and time formats correctly preserved.",
                category: "Punctuation / numbers\nSimple sentence\nS\nRobustness validation"
            }
        ];

        this.dataList.push(...positiveCases);
        console.log(`Added ${positiveCases.length} positive cases`);
    }

    setupNegativeData() {
        console.log("Setting up negative functional data...");

        const negativeCases = [
            {
                id: "N_TC_01",
                name: "Short abbreviation test",
                length: "S",
                input: "Thx machan!",
                expected: "Thx මචන්!",
                justification: "Chat shorthand 'Thx' not recognized by system.",
                category: "Slang / informal language\nSimple sentence\nS\nRobustness validation"
            },
            {
                id: "N_TC_02",
                name: "Unusual slang test",
                length: "M",
                input: "adoo vaedak baaragaththaanam eeka hariyata karapanko bn",
                expected: "අඩෝ වැඩක් බාරගත්තානම් එක හරියට කරපන්කො බන්",
                justification: "Very informal slang 'bn' may not convert correctly.",
                category: "Slang / informal language\nComplex sentence\nM\nRobustness validation"
            },
            {
                id: "N_TC_03",
                name: "Extreme concatenation test",
                length: "S",
                input: "mamagedharayanavaamatapaankannaooencehetaapiyanawa",
                expected: "මමගෙදරයනවාමටපාන්කන්නඕඑනcඑහෙටාපියනවා",
                justification: "Extremely joined words without any spaces may cause parsing errors.",
                category: "Formatting (spaces / line breaks / paragraph)\nSimple sentence\nS\nRobustness validation"
            },
            {
                id: "N_TC_04",
                name: "Excessive spacing test",
                length: "S",
                input: "mama   gedhara   yanavaa.",
                expected: "මම ගෙදර යනවා.",
                justification: "Multiple spaces between words should ideally be collapsed.",
                category: "Formatting (spaces / line breaks / paragraph)\nSimple sentence\nS\nRobustness validation"
            },
            {
                id: "N_TC_05",
                name: "Mixed case abbreviation test",
                length: "S",
                input: "Mata CPU eka replace karanna onne",
                expected: "මට CPU එක replace කරන්න ඕනෙ.",
                justification: "Mixed case Singlish 'onne' instead of 'oone' edge case.",
                category: "Mixed Singlish + English\nSimple sentence\nS\nRobustness validation"
            },
            {
                id: "N_TC_06",
                name: "Punctuation spam test",
                length: "S",
                input: "ehema karanna pluwandha??? !!!",
                expected: "එහෙම කරන්න පුළුවන්ද??? !!!",
                justification: "Multiple punctuation marks edge case.",
                category: "Punctuation / numbers\nSimple sentence\nS\nRobustness validation"
            },
            {
                id: "N_TC_07",
                name: "Currency symbol test",
                length: "S",
                input: "Rs. 5343 denna",
                expected: "Rs. 5343 දෙන්න",
                justification: "Currency symbol followed by number edge case.",
                category: "Punctuation / numbers\nImperative (command)\nS\nRobustness validation"
            },
            {
                id: "N_TC_08",
                name: "Long mixed content test",
                length: "L",
                input: "mema kramayen katayuthu kiriima mata sudhusudha kiyalaa karuNaakara dhaenum dhenna. mee vidhihata kaLoth kaarYAya saralava saha kaalaya ithiri karagena sampuurNa karanna puLuvan kiyalaa mata hithenavaa. ee nisaa mee kramaya Bhaavithaa kaLaata prashnayak thiyenavadha kiyalaa karuNaakara kiyanna. mama ithin dhaen karanna hadanavaa api heta sathiyee kalin yanna hadanavaa oyaata kohomadha heta yanna puluvandha kiyalaa danaganna oone.",
                expected: "මෙම ක්‍රමයෙන් කටයුතු කිරීම මට සුදුසුද කියලා කරුණාකර දැනුම් දෙන්න. මේ විදිහට කළොත් කාර්යය සරලව සහ කාලය ඉතිරි කරගෙන සම්පූර්ණ කරන්න පුළුවන් කියලා මට හිතෙනවා. ඒ නිසා මේ ක්‍රමය භාවිතා කළාට ප්‍රශ්නයක් තියෙනවද කියලා කරුණාකර කියන්න. මම ඉතින් දැන් කරන්න හදනවා අපි හෙට සතියේ කලින් යන්න හදනවා ඔයාට කොහොමද හෙට යන්න පුළුවන්ද කියලා දැනගන්න ඕනෙ.",
                justification: "Extremely long input stress test.",
                category: "Formatting (spaces / line breaks / paragraph)\nComplex sentence\nL\nRobustness validation"
            },
            {
                id: "N_TC_09",
                name: "Rare colloquialism test",
                length: "S",
                input: "eka poddak amaaru wedak vagee",
                expected: "එක පොඩ්ඩක් අමාරු වැඩක් වගේ.",
                justification: "Regional colloquial variation edge case.",
                category: "Slang / informal language\nSimple sentence\nS\nRobustness validation"
            },
            {
                id: "N_TC_10",
                name: "Tense inconsistency test",
                length: "M",
                input: "mama iye giyaa, dhaen kanawa, heta enavaa.",
                expected: "මම ඉයෙ ගියා, දැන් කනවා, හෙට එනවා.",
                justification: "Rapid tense switching edge case.",
                category: "Daily language usage\nCompound sentence\nM\nRobustness validation"
            }
        ];

        this.dataList.push(...negativeCases);
        console.log(`Added ${negativeCases.length} negative cases`);
    }

    setupUIData() {
        console.log("Setting up UI test data...");

        const uiCases = [
            {
                id: "UI_TC_01",
                name: "Real-time update test",
                length: "S",
                input: "mama gedhara yanavaa",
                expected: "මම ගෙදර යනවා",
                justification: "Sinhala output should update automatically while typing.",
                category: "Usability flow (real-time conversion)\nSimple sentence\nS\nReal-time output update behavior"
            },
            {
                id: "UI_TC_02",
                name: "Clear button test",
                length: "S",
                input: "mama gedhara yanavaa",
                expected: "",
                justification: "Clear button should reset both fields.",
                category: "Empty/cleared input handling\nSimple sentence\nS\nError handling / input validation"
            }
        ];

        this.dataList.push(...uiCases);
        console.log(`Added ${uiCases.length} UI data cases`);
    }

    saveToExcel(filePath) {
        console.log(`Saving data to ${filePath}...`);

        const workbook = utils.book_new();

        const testCaseData = [
            ["TC ID", "Test case name", "Input length type", "Input", "Expected output", "Actual output", "Status", "Accuracy justification/ Description of issue type", "What is covered by the test"]
        ];

        this.dataList.forEach(tc => {
            testCaseData.push([
                tc.id,
                tc.name,
                tc.length,
                tc.input,
                tc.expected,
                "",
                "",
                tc.justification || "",
                tc.category
            ]);
        });

        const testCaseSheet = utils.aoa_to_sheet(testCaseData);
        const wscols = [
            { wch: 12 }, { wch: 40 }, { wch: 15 }, { wch: 50 }, { wch: 50 }, { wch: 50 }, { wch: 8 }, { wch: 60 }, { wch: 40 }
        ];
        testCaseSheet['!cols'] = wscols;
        utils.book_append_sheet(workbook, testCaseSheet, "Test cases");

        const conventionData = [
            ["Test Case ID Formats:"],
            ["P_TC_xx for Positive Functional"],
            ["N_TC_xx for Negative Functional"],
            ["UI_TC_xx for UI/Usability"],
            [""],
            ["Length codes:"],
            ["S: Short, M: Medium, L: Long"]
        ];
        utils.book_append_sheet(workbook, utils.aoa_to_sheet(conventionData), "Conventions");

        const summaryData = [
            ["Test Coverage Summary"],
            [""],
            ["Total Test Cases:", this.dataList.length],
            ["Positive:", this.dataList.filter(tc => tc.id.startsWith('P_TC')).length],
            ["Negative:", this.dataList.filter(tc => tc.id.startsWith('N_TC')).length],
            ["UI Tests:", this.dataList.filter(tc => tc.id.startsWith('UI_TC')).length]
        ];
        utils.book_append_sheet(workbook, utils.aoa_to_sheet(summaryData), "Coverage");

        const executionData = [
            ["Testing Guide"],
            ["1. Install: npm install"],
            ["2. Run Tests: node src/tests/test-runner.js"],
            [""],
            ["Output paths:"],
            ["- results/test-results.xlsx"],
            ["- results/execution-report.html"]
        ];
        utils.book_append_sheet(workbook, utils.aoa_to_sheet(executionData), "How to Run");

        const dir = dirname(filePath);
        if (!existsSync(dir)) {
            mkdirSync(dir, { recursive: true });
        }

        writeFile(workbook, filePath);
        console.log(`Data saved to: ${filePath}`);

        console.log("\nDATA SETUP COMPLETE:");
        console.log("==============================");
        console.log(`Total Cases: ${this.dataList.length}`);
        console.log(`Positive: ${this.dataList.filter(tc => tc.id.startsWith('P_TC')).length}/24`);
        console.log(`Negative: ${this.dataList.filter(tc => tc.id.startsWith('N_TC')).length}/10`);
        console.log(`UI Tests: ${this.dataList.filter(tc => tc.id.startsWith('UI_TC')).length}/2`);
    }
}

if (require.main === module) {
    console.log("Starting Data Setup...");
    try {
        const manager = new TestCaseManager();
        manager.collectAllData();
        manager.saveToExcel("test-data/test-cases.xlsx");
        console.log("\nSetup complete!");
    } catch (error) {
        console.error("Error:", error.message);
        process.exit(1);
    }
}

module.exports = TestCaseManager;