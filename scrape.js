function rectDistance(a, b) {
    let aRight = a.offsetLeft + a.clientWidth;
    let bRight = b.offsetLeft + b.clientWidth;
    let xDist = Math.max(0, Math.max(b.offsetLeft-aRight, a.offsetLeft-bRight));
    let aBottom = a.offsetTop + a.clientHeight;
    let bBottom = b.offsetTop + b.clientHeight;
    let yDist = Math.max(0, Math.max(b.offsetTop-aBottom, a.offsetTop-bBottom));
    return Math.max(xDist, yDist);
}

let hasText = function(n){return n.innerText != "";};
window.divs = Array.prototype.slice.call(window.frames[3].document.querySelectorAll("#divHtmlReport div[class^=mls]")).filter(hasText);

for(let i=0; i < window.divs.length-1; i++) {
    let di = window.divs[i];
    di.divIdx = i;
    for(let j=i+1; j < window.divs.length; j++) {
        let dj = window.divs[j];
        let i_x = di.offsetLeft, i_y = di.offsetTop;
        let j_x = dj.offsetLeft, j_y = dj.offsetTop;
        let dist = rectDistance(di, dj);
        let i_direction = "", j_direction = "";
        if (i_y !== j_y) {
            const verticalDirection = i_y < j_y ? ["lower", "upper"] : ["upper", "lower"];
            [i_direction, j_direction] = verticalDirection;
        }
        if (i_x !== j_x) {
            const horizontalDirection = i_x < j_x ? ["right", "left"] : ["left", "right"];
            i_direction += horizontalDirection[0];
            j_direction += horizontalDirection[1]
        }
        if (i_direction == "" || j_direction == "") continue;
        if (!di.hasOwnProperty(i_direction) || rectDistance(di[i_direction], di) > dist) {
            di[i_direction] = dj;
        }
        if (!dj.hasOwnProperty(j_direction) || rectDistance(dj[j_direction], dj) > dist) {
            dj[j_direction] = di;
        }
    }
}

var data = [
    ["Listing ID"],
    ["Address #1"],
    ["Address #2"],
    ["Address #3"],
    ["Address #4"],
    ["Listing Price"],
    ["Gross Taxes"],
    ["Maintenance Fee"],
    ["Number of Bedrooms"],
    ["Number of Bathrooms"],
    ["Total SQFT"],
    ["Listing Description"]
];

let nodeWithText = function(t){return divs.find(function(i){return i.innerText.includes(t)})};

let status = divs.find(function(i){return !i.hasOwnProperty("left") && !i.hasOwnProperty("upperleft") && !i.hasOwnProperty("upper") && !i.hasOwnProperty("upperright")})
let listingId = status ? status.lower : null;
let addr1 = status ? status.right : null;
let addr2 = addr1 ? addr1.lower : null;
let addr3 = addr2 ? addr2.lower : null;
let addr4 = addr3 ? addr3.lower : null;
let lp = nodeWithText("(LP)");
let listingPrice = lp ? lp.left : null;
let gt = nodeWithText("Gross Taxes");
let grossTaxes = gt ? gt.right : null;
let bed = nodeWithText("Bedrooms:");
let bedrooms = bed ? bed.right : null;
let bath = nodeWithText("Bathrooms:");
let bathrooms = bath ? bath.right : null;
let sf = nodeWithText("Grand Total:");
let sqft = sf ? sf.right : null;
let description = divs.slice().sort(function(a,b){ return b.innerText.length - a.innerText.length })[0];

function addColumn(columnNo, node) {
    data[columnNo].push(node ? node.innerText : "");
}

addColumn(0, listingId);
addColumn(1, addr1);
addColumn(2, addr2);
addColumn(3, addr3);
addColumn(4, addr4);
addColumn(5, listingPrice);
addColumn(6, grossTaxes);
addColumn(7, null);
addColumn(8, bedrooms);
addColumn(9, bathrooms);
addColumn(10, sqft);
addColumn(11, description);

/* create worksheet */
var ws = XLSX.utils.json_to_sheet(data, {});

/* create workbook and export */
var wb = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wb, ws, 'MLS');
XLSX.writeFile(wb, "paragonrels.xlsx");
