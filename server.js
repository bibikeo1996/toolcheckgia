var express =   require("express");
var multer  =   require('multer');
const puppeteer = require('puppeteer');
var excel = require('excel4node');
var app     =   express();
app.use(express.static(__dirname))
//var port = process.env.PORT || 3000;
function expxortLink(data){
	var workbook = new excel.Workbook();
	var worksheet = workbook.addWorksheet('DMCL');
	const headings = ['Link'];
	headings.forEach((heading, index) => {
        worksheet.cell(1, index + 1).string(heading);
    })
    data.forEach((item, index) => {
    	worksheet.cell(index + 2, 1).string(item.Link);
    });
    var today1 = new Date();
    var filename1 = "LinkDMCL"+ '-' + today1.getDate() + '-' +(today1.getMonth()+1) + '-' + today1.getFullYear() +'.xlsx';
    workbook.write(filename1);
    return filename1;
}

function exportToExcelDoiThu(data) {
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();

    // Add Worksheets to the workbook
    var worksheet = workbook.addWorksheet('Sheet 1');

    const headings = ['Tên', 'Giá Thường',"Giá KM", 'Buộc trừ 1', 'Buộc trừ 2','Trừ tiền 1', 'Trừ tiền 2', 'Tổng Giá Quà'];//, 'Gift6', 'Gift7'];;
    worksheet.row(1).freeze();
    // Writing from cell A1 to I1
    headings.forEach((heading, index) => {
        worksheet.cell(1, index + 1).string(heading);
    })

    // Writing from cell A2 to I2 , A3 to I3, .....
    data.forEach((item, index) => {
        worksheet.cell(index + 2, 1).string(item.Name);
        worksheet.cell(index + 2, 2).string(item.Price);
        worksheet.cell(index + 2, 3).string(item.Price2);
        worksheet.cell(index + 2, 4).string(item.Gift);
        worksheet.cell(index + 2, 5).string(item.Gift2);
        worksheet.cell(index + 2, 6).string(item.Gift3);
        worksheet.cell(index + 2, 7).string(item.Gift4);
        worksheet.cell(index + 2, 8).string(item.Tonggiatri);
        // worksheet.cell(index + 2, 8).string(item.Gift6);
        // worksheet.cell(index + 2, 9).string(item.Gift7);
    });
    var today1 = new Date();
    var filename = "GiaDMX-NK"  + '-' + today1.getDate() + '-' +(today1.getMonth()+1) + '-' + today1.getFullYear() + '.xlsx'; //+ Date.now().toString()
    workbook.write(filename);
    return filename;
}
function exportToExcel(data) {
    // Create a new instance of a Workbook class
    var workbook = new excel.Workbook();
    // Add Worksheets to the workbook
    var worksheet = workbook.addWorksheet('DMCL');
    var style = workbook.createStyle({
      font: {
        size: 12
      },
      alignment: {
        horizontal: 'center'
      },
    });

    worksheet.row(1).freeze();
    worksheet.column(7).hide();
    worksheet.column(8).hide();
    worksheet.column(9).hide();
    worksheet.column(10).hide();
    worksheet.column(11).hide();
    worksheet.column(12).hide();
    const headings = ['SAP', 'Model', 'NY','Giá Cuối','Note','Loại','Góp 0%',"Bài Viết","Slide","Sticker","Icon Giảm Thêm","Icon Big Sale","Quà","Quà 2","LayOut","Format","Giảm miệng","%"]//,"LoạiSP","LoạiSP2"];//, 'Buộc trừ 2','Trừ tiền 1', 'Trừ tiền 2', 'Tổng Giá Quà'];//, 'Gift6', 'Gift7'];
    //const headings = ['Article', 'Cmt', 'Loại'];
    // Writing from cell A1 to I1
    headings.forEach((heading, index) => {
        worksheet.cell(1, index + 1).string(heading).style(style);
    })
    
    // Writing from cell A2 to I2 , A3 to I3, .....
    data.forEach((item, index) => {
        worksheet.cell(index + 2, 1).string(item.Sap);
        worksheet.cell(index + 2, 2).string(item.Name);
        worksheet.cell(index + 2, 3).string(item.Price1);
       	worksheet.cell(index + 2, 4).string(item.Price2);
        worksheet.cell(index + 2, 5).string(item.Comment);
        worksheet.cell(index + 2, 6).string(item.Loai);
        worksheet.cell(index + 2, 7).string(item.Gop);
        worksheet.cell(index + 2, 8).string(item.Baiviet);
        worksheet.cell(index + 2, 9).string(item.Slide);
        worksheet.cell(index + 2, 10).string(item.Sticker);
        worksheet.cell(index + 2, 11).string(item.Giamthem);
        worksheet.cell(index + 2, 12).string(item.SaleBig);
        worksheet.cell(index + 2, 13).string(item.Gift);
        worksheet.cell(index + 2, 14).string(item.Gift2);
        worksheet.cell(index + 2, 15).string(item.LayOut);
        worksheet.cell(index + 2, 16).formula('IF(AND(C2=D2),"Show",IF(AND(C2>D2),"Giấu","Sai"))');
        worksheet.cell(index + 2, 17).formula('C2-D2');
        worksheet.cell(index + 2, 18).formula('ROUND((Q2*100)/C2,0)');
        // worksheet.cell(index + 2, 15).string(item.LoaiSP2);
    });
    var today1 = new Date();
    var filename = "GiaDMCL"  + '-' + today1.getDate() + '-' +(today1.getMonth()+1) + '-' + today1.getFullYear() + '.xlsx'; //+ Date.now().toString()
    workbook.write(filename);
    return filename;
}

function getLink(file){
    const XLSX = require('xlsx');
    var workbook = XLSX.readFile(file);
    var sheet_name_list = workbook.SheetNames;
    var urls = [];
    data = XLSX.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]); 
    return data;
};

var storage =   multer.diskStorage({
  destination: function (req, file, callback) {
    callback(null, __dirname);
  },
  filename: function (req, file, callback) {
    callback(null,file.originalname);
  }
});


var upload = multer({ storage : storage});

app.get('/',function(req,res){
      res.sendFile(__dirname + "/views/index.html");
});

// ================================================ Code Check Giá DMCL Bằng Article ================================================

app.post('/checkgiadmcl',upload.single('checkgiadmcl'),function(req,res){
	if(req === null){
		res.write("<p>Error!</p>")
	}else{
        	var file = req.file.originalname;
        	res.write("<p>Checking...</p>");
        	res.write("<p>"+file+"</p>");
        	(async () => {
			    const browser = await puppeteer.launch({
					  'args' : [
					    '--no-sandbox',
					    '--disable-setuid-sandbox'
					  ]
					});
			    const page = await browser.newPage()
			    await page.goto('https://dienmaycholon.vn/admin/users/login',{waitUntil: 'load', timeout:300000});
			    let user = await page.$('#username');
			    let passw = await page.$('#passwords');
			    await user.type("thanhliem");
			    await passw.type("beo4356143");
			    // await page.type('#username', 'thanhliem');// tài khoản admin
			    // await page.type('#passwords', 'beo4356143');// mật khẩu admin

			    await page.click('.button');
			    const urls = getLink(file);
			    await page.setRequestInterception(true);
			            page.on('request', (req) => {
			                if(req.resourceType() == 'stylesheet' || req.resourceType() == 'font' || req.resourceType() == 'image' || req.resourceType() ==
			                 'script'){
			                    req.abort();
			                }
			                else {
			                    req.continue();
			                }
			            });
			    let arrInfo = [];
			    let result = [];
			    for (var rs of urls) { // open for 1
			        try { // open try 1
			            await page.goto("https://dienmaycholon.vn/tu-khoa/"+rs.Article,{waitUntil: 'load', timeout: 100000});
			            const info1 = await page.evaluate(() => {
			                let checkweb = document.querySelector(".item_product")
			                if (checkweb !== null) {
			                   const checklink = document.querySelector(".item_product .pro_infomation a")
			                   const link = document.querySelector(checklink !== null ? ".item_product .pro_infomation a" : ".khongcoclass")
			                       return {
			                            //...data,
			                            Link: link ? "https://dienmaycholon.vn"+link.getAttribute('href') : "Not found"
			                        }
			                }
			                return {
			                    Link: "https://dienmaycholon.vn/tu-khoa/khong-tim-thay"
			                };
			            })
			            if(info1){
			                result.push(info1)
			                res.write("<p>" + rs.Article + " => Xong! </p>") 
			            }
			        } // end try 1
			        catch (err) {
			            res.write("<p>Error => " + rs.Article +"</p>")
			        }
			    } // end for 1
			    var tenlink = expxortLink(result)
			    expxortLink(result);
			    res.write("<a href='/"+tenlink+"'>Download File Link</a>");
			    for(var rs2 of result){ // open for check giá 
			    	try{
			    		await page.goto(rs2.Link, {waitUntil: 'load', timeout: 300000});
			    		const info = await page.evaluate(() => { // open page evaluate 
			    			let checkweb = document.querySelector(".big_detail")
			    			if(checkweb !== null){ // open if check giá dmcl
			    				const checkname = document.querySelector("h1.product_name")
			                    const checksap = document.querySelector("strong.sapcode")
			                    const checkgop = document.querySelector("i.iconsprites_payment")
			                    const checkgop2 = document.querySelector("i.iconsprites_payment_0d")
			                    const checkbaiviet = document.querySelector(".area_article h3")
			                    const checkbaiviet1 = document.querySelector(".area_article p")
			                    const checkslide = document.querySelector(".owl-stage .desc_picture")
			                    const checklayout = document.querySelector(".box_price_layout_cost")
			                    const checksticker = document.querySelector(".status_info div img")
			                    const checkicongiamthem = document.querySelector(".product_seeing .mobisprites_centernew")
			                    const checkiconsalebig = document.querySelector(".iconsalesalebig")

			                    function checkGiaNiemYet()
			                                {
			                                    const check = document.querySelectorAll(".box_price_layout_cost strong, strong.price_sale");
			                                    const done = [];
			                                    for (var i = 0; i < check.length; i++) {
			                                        const data = check[i].innerText.toLowerCase();
			                                      if (data.includes(".000đ"))
			                                      {
			                                        const gia = /[0-9]/g;
			                                        var first = data.search(gia);
			                                        var last = data.search("đ");
			                                        const rs = data.slice(first,last);
			                                        done.push(rs.replace(/[^A-Z0-9]/ig,''));
			                                        break;
			                                      }
			                                    }return done; // trả về mảng dữ liệu
			                                }
			                    function checkGiaCuoi()
			                                {
			                                    const check = document.querySelectorAll(".detail_more");
			                                    const done = [];
			                                    for (var i = 0; i < check.length; i++) {
			                                        const data = check[i].innerText.toLowerCase();
			                                      if (data.includes("giá data đối chiếu"))
			                                      {
			                                        const kytu = /[:]/gi;
			                                        var first = data.search(kytu);
			                                        const rs = data.slice(first,49);
			                                        done.push(rs.replace(/[^A-Z0-9]/ig,''));
			                                        break;
			                                      }
			                                    }return done; // trả về mảng dữ liệu
			                                }
			                    function checkLoaiSP()
			                                {
			                                    const check = document.querySelectorAll(".detail_more");
			                                    const done = [];
			                                    for (var i = 0; i < check.length; i++) {
			                                        const data = check[i].innerText.toLowerCase();
			                                      if (data.includes("giá data đối chiếu"))
			                                      {
			                                        const rs = data.slice(5,8).toUpperCase();  // cắt chuỗi bắt đầu từ ký tự thứ 5 đến 8
			                                        done.push(rs.replace(/ /g,''));
			                                        break;
			                                      }
			                                    }return done; // trả về mảng dữ liệu
			                                }
			                    function checkComment()
			                                {
			                                    const checkcomment = document.querySelectorAll(".detail_more");
			                                    const done = [];
			                                    for(var i = 0; i < checkcomment.length; i++){
			                                        const data = checkcomment[i].innerHTML.toLowerCase();
			                                        if(data.includes("giá online") || data.includes("cng:") || data.includes("giá cuối") || data.includes("đã giảm")){
			                                            done.push(data);break;
			                                        }else if(data.includes("khuyến mãi")){
			                                        	done.push("Khuyến Mãi");break;
			                                        }
			                                    }return done;
			                                }
			                    function checkQua()
			                                {
			                                    const checkqua = document.querySelectorAll("ul.main_gift li a");
			                                    const done = [];
			                                    for(var i = 0; i < checkqua.length; i++){
			                                        const data = checkqua[i].innerHTML.toLowerCase();
			                                        if(data.includes("sale khổng lồ")
			                                        	|| data.includes("thanh toán qua thẻ giảm thêm") 
			                                        	|| data.includes("giảm thêm")
			                                        	|| data.includes("rẻ ngỡ ngàng rẻ hơn đến")
			                                        	|| data.includes("gọi hotline giảm thêm") 
			                                        	|| data.includes("sale hàng hiệu") 
			                                        	|| data.includes("mua online giảm thêm") 
			                                        	|| data.includes("sale hàng hiệu giảm thêm") 
			                                        	|| data.includes("gọi hotline") 
			                                        	|| data.includes("đã giảm")){
			                                            done.push(data);
			                                        }
			                                    }return done;
			                                }
			    				
			    				const gianiemyet = checkGiaNiemYet()
			                    const giacuoi = checkGiaCuoi()
			                    const sap = document.querySelector(checksap !== null ? "strong.sapcode" : ".khongcoclass")
			                    const name = document.querySelector(checkname !== null ? "h1.product_name" : ".khongcoclass")
			                    const gop = (checkgop !== null ? "Có" : checkgop2 !== null ? "Có" : "Không")
			                    const baiviet = (checkbaiviet !== null ? "Có" : "Không")
			                    const sticker = (checksticker !== null ? "Có" : "Không")
			                    const layout = (checklayout !== null ? "Có" : "Không")
			                    const slide = (checkslide !== null ? "Có" : "Không")
			                    const icongiamthem = (checkicongiamthem !== null ? "Có" :"Không")
			                    const iconsalebig = (checkiconsalebig !== null ? "Có" :"Không")
			                    const loaisp = checkLoaiSP()
			                    const comment = checkComment()
			                    const laygift = checkQua()
			                    const gift = laygift[0];
			                    const gift2 = laygift[1];
			                    let data = {
			                        Name: name ? name.innerText : "Not found",
			                        Sap: sap ? sap.innerText : "Not found",
			                        Price1: gianiemyet ? gianiemyet : "Not found",
			                        Price2: giacuoi ? giacuoi : "Not found",
			                        Comment: comment ? comment : "Not found",
			                        Loai: loaisp ? loaisp : "Not found",
			                        Gop: gop ? gop : "Not found",
			                        Baiviet: baiviet ? baiviet : "Not found",
			                        Slide: slide ? slide : "Not found",
			                        Sticker: sticker ? sticker : "Not found",
			                        Giamthem: icongiamthem ? icongiamthem : "Not found",
			                        SaleBig: iconsalebig ? iconsalebig : "Not found",
			                        Gift: gift ? gift : "Not found",
			                        Gift2: gift2 ? gift2 : "Not found",
			                        LayOut: layout ? layout : "Not found",
			                    }
			                    return {
			                           ...data,
			                        }
			                    }
			                return {
			                    Name: "Sai!",
			                    Sap: "Sai!",
			                    Price1: "Sai!",
			                    Price2: "Sai!",
			                    Comment: "Sai!",
			                    Loai:"Sai!",
			                    Gop:"Sai!",
			                    Baiviet:"Sai!",
			                    Slide:"Sai!",
			                    Sticker:"Sai!",
			                    Giamthem:"Sai!",
			                    SaleBig:"Sai!",
			                    Gift:"Sai!",
			                    Gift2:"Sai!",
			                    LayOut:"Sai!",
			    			} // close if check giá dmcl
			    		}) // end page evaluate 
						 if(info){
			            	arrInfo.push(info)
			                res.write("<p>" + info.Sap + " => Xong!</p>")
			            }
			    	} catch (err) {
			    		res.write("<p>Error! => " + rs2.Link + "</p>")
			    	}
			    }
			    var tenlink = exportToExcel(arrInfo)
			    exportToExcel(arrInfo);
			    res.write("<a href='/'>Back</a></br>");
			    res.end("<a href='/"+tenlink+"'>Download File</a>");
			})(); // end async check gia dmcl
		} //end else
        }); // end app.post

// ================================================ Code Check Giá DMCL Bằng Link ================================================


app.post('/checkbanglink',upload.single('checkbanglink'),function(req,res){ // open app.post 2
	if(req === null){
		 return res.write("Error uploading file.");
		 console.log(err)
		}else{
        	var file = req.file.originalname;
        	res.write("<p>Checking...</p>");
    		(async () => { // open async 2
					const browser = await puppeteer.launch({
					  'args' : [
					    '--no-sandbox',
					    '--disable-setuid-sandbox'
					  ]
					});

			    const page = await browser.newPage()
			    await page.goto('https://dienmaycholon.vn/admin/users/login',{waitUntil: 'load', timeout:100000});

			   let user = await page.$('#username');
			    let passw = await page.$('#passwords');
			    await user.type("thanhliem");
			    await passw.type("beo4356143");

			    await page.click('.button');
			    const urls = getLink(file);
			    await page.setRequestInterception(true);
		        page.on('request', (req) => {
		                if(req.resourceType() == 'stylesheet' || req.resourceType() == 'font' || req.resourceType() == 'image' || req.resourceType() ==
		                 'script'){
		                    req.abort();
		                }
		                else {
		                    req.continue();
		                }
		            });
		    	let arrInfo = [];
		    	for(var rs of urls){ // open for 2
		    		try{
		    			await page.goto(rs.Link, {waitUntil: 'load', timeout: 300000});
		    			const info = await page.evaluate(() => { //open evaluate 2
		    			let checkweb = document.querySelector(".big_detail")
		    			if(checkweb !== null){ // open if
		    				const checkname = document.querySelector("h1.product_name")
	                    const checksap = document.querySelector("strong.sapcode")
	                    const checkgop = document.querySelector("i.iconsprites_payment")
	                    const checkgop2 = document.querySelector("i.iconsprites_payment_0d")
	                    const checkbaiviet = document.querySelector(".area_article h3")
	                    const checkbaiviet1 = document.querySelector(".area_article p")
	                    const checkslide = document.querySelector(".owl-stage .desc_picture")
	                    const checklayout = document.querySelector(".box_price_layout_cost")
	                    const checksticker = document.querySelector(".status_info div img")
	                    const checkicongiamthem = document.querySelector(".product_seeing .mobisprites_centernew")
	                    const checkiconsalebig = document.querySelector(".iconsalesalebig")
	                    function checkGiaNiemYet()
			                                {
			                                    const check = document.querySelectorAll(".box_price_layout_cost strong, strong.price_sale");
			                                    const done = [];
			                                    for (var i = 0; i < check.length; i++) {
			                                        const data = check[i].innerText.toLowerCase();
			                                      if (data.includes(".000đ"))
			                                      {
			                                        const gia = /[0-9]/g;
			                                        var first = data.search(gia);
			                                        var last = data.search("đ");
			                                        const rs = data.slice(first,last);
			                                        done.push(rs.replace(/[^A-Z0-9]/ig,''));
			                                        break;
			                                      }
			                                    }return done; // trả về mảng dữ liệu
			                                }
			            function checkGiaCuoi()
			                                {
			                                    const check = document.querySelectorAll(".detail_more");
			                                    const done = [];
			                                    for (var i = 0; i < check.length; i++) {
			                                        const data = check[i].innerText.toLowerCase();
			                                      if (data.includes("giá data đối chiếu"))
			                                      {
			                                        const kytu = /[:]/gi;
			                                        var first = data.search(kytu);
			                                        const rs = data.slice(first,49);
			                                        done.push(rs.replace(/[^A-Z0-9]/ig,''));
			                                        break;
			                                      }
			                                    }return done; // trả về mảng dữ liệu
			                                }
			            function checkLoaiSP()
			                                {
			                                    const check = document.querySelectorAll(".detail_more");
			                                    const done = [];
			                                    for (var i = 0; i < check.length; i++) {
			                                        const data = check[i].innerText.toLowerCase();
			                                      if (data.includes("giá data đối chiếu"))
			                                      {
			                                        const rs = data.slice(5,8).toUpperCase();  // cắt chuỗi bắt đầu từ ký tự thứ 5 đến 8
			                                        done.push(rs.replace(/ /g,''));
			                                        break;
			                                      }
			                                    }return done; // trả về mảng dữ liệu
			                                }
			            function checkComment()
			                                {
			                                    const checkcomment = document.querySelectorAll(".detail_more");
			                                    const done = [];
			                                    for(var i = 0; i < checkcomment.length; i++){
			                                        const data = checkcomment[i].innerHTML.toLowerCase();
			                                        if(data.includes("giá online") || data.includes("cng:") || data.includes("giá cuối") || data.includes("đã giảm")){
			                                            done.push(data);break;
			                                        }else if(data.includes("khuyến mãi")){
			                                        	done.push("Khuyến Mãi");break;
			                                        }
			                                    }return done;
			                                }
			            function checkQua()
			                                {
			                                    const checkqua = document.querySelectorAll("ul.main_gift li a");
			                                    const done = [];
			                                    for(var i = 0; i < checkqua.length; i++){
			                                        const data = checkqua[i].innerHTML.toLowerCase();
			                                        if(data.includes("sale khổng lồ")
			                                        	|| data.includes("thanh toán qua thẻ giảm thêm") 
			                                        	|| data.includes("giảm thêm")
			                                        	|| data.includes("rẻ ngỡ ngàng rẻ hơn đến")
			                                        	|| data.includes("gọi hotline giảm thêm") 
			                                        	|| data.includes("sale hàng hiệu") 
			                                        	|| data.includes("mua online giảm thêm") 
			                                        	|| data.includes("sale hàng hiệu giảm thêm") 
			                                        	|| data.includes("gọi hotline") 
			                                        	|| data.includes("đã giảm")){
			                                            done.push(data);
			                                        }
			                                    }return done;
			                                }
			    				
			    		const gianiemyet = checkGiaNiemYet()
			            const giacuoi = checkGiaCuoi()
			            const sap = document.querySelector(checksap !== null ? "strong.sapcode" : ".khongcoclass")
			            const name = document.querySelector(checkname !== null ? "h1.product_name" : ".khongcoclass")
			            const gop = (checkgop !== null ? "Có" : checkgop2 !== null ? "Có" : "Không")
			            const baiviet = (checkbaiviet !== null ? "Có" : "Không")
			            const sticker = (checksticker !== null ? "Có" : "Không")
			            const layout = (checklayout !== null ? "Có" : "Không")
			            const slide = (checkslide !== null ? "Có" : "Không")
			            const icongiamthem = (checkicongiamthem !== null ? "Có" :"Không")
			            const iconsalebig = (checkiconsalebig !== null ? "Có" :"Không")
			            const loaisp = checkLoaiSP()
			            const comment = checkComment()
			            const laygift = checkQua()
			            const gift = laygift[0];
			            const gift2 = laygift[1];
			            let data = {
                        Name: name ? name.innerText : "Not found",
                        Sap: sap ? sap.innerText : "Not found",
                        Price1: gianiemyet ? gianiemyet : "Not found",
                        Price2: giacuoi ? giacuoi : "Not found",
                        Comment: comment ? comment : "Not found",
                        Loai: loaisp ? loaisp : "Not found",
                        Gop: gop ? gop : "Not found",
                        Baiviet: baiviet ? baiviet : "Not found",
                        Slide: slide ? slide : "Not found",
                        Sticker: sticker ? sticker : "Not found",
                        Giamthem: icongiamthem ? icongiamthem : "Not found",
                        SaleBig: iconsalebig ? iconsalebig : "Not found",
                        Gift: gift ? gift : "Not found",
                        Gift2: gift2 ? gift2 : "Not found",
                        LayOut: layout ? layout : "Not found",
	                    }
	                    return {
	                           ...data
	                        }
		    			}else{
		    				res.write("<p>Error!</p>")
		    			}//end if
	                }) // end evaluate 2
					if(info){
			                arrInfo.push(info)
			                res.write("<p>Stt " + rs.STT + " => " + info.Sap + " => Done!</p>")
			            }
		    		} catch (err){
		    			 res.write("<p>Errors => " + rs.STT +"</p>")
		    		}
		    	} // end for 2
		    	var tenlink = exportToExcel(arrInfo)
			    exportToExcel(arrInfo);
			    res.write("<a href='/'>Back</a></br>");
			    res.end("<a href='/"+tenlink+"'>Download File</a>");
    		})(); // end async 2
    	}
    	 }); // end app post

// ================================================ Code Check Giá DMX/NK ================================================

app.post('/checkgiankdmx',upload.single('checkgiankdmx'),function(req,res){
	if(req === null){
		res.write("<p>Errors!!</p>")
	}else{
	var file = req.file.originalname;
	res.write("<p>Checking...</p>");
	(async () => { 
	const browser = await puppeteer.launch({
					  'args' : [
					    '--no-sandbox',
					    '--disable-setuid-sandbox'
					  ]
					});
    const page = await browser.newPage()
    const urls = getLink(file);
    console.log(urls);
    await page.setRequestInterception(true);
            page.on('request', (req) => {
                if(req.resourceType() == 'stylesheet' || req.resourceType() == 'font' || req.resourceType() == 'image' || req.resourceType() ==
                 'script'){
                    req.abort();
                }
                else {
                    req.continue();
                }
            });
    let arrInfo = [];
    for (var doithu of urls) {
        try {
            await page.goto(doithu.Link, {waitUntil: 'load', timeout: 100000});
            const info = await page.evaluate(() => {
                let checkweb1 = document.querySelector("#main-container")
                let checkweb2 = document.querySelector(".NkPdp_productInfo")
                if (checkweb1 !== null || checkweb2 !== null) {
                    const checknamedmx = document.querySelector("h1");
                    const checknamenk = document.querySelector("h1.product_info_name")
                    const checktonggiatri = document.querySelector(".area_promotion strong b")
                    const checktonggiatri1 = document.querySelector(".boxshockheader")
                    function checkGia()
                    {
                    	const check = document.querySelectorAll(".displayp strong, .kmgiagach, span.five-7ngay, .nosell, .no-sell strong, .area_order b, .box-info strong, .price_shock_online_exp, .nk-price-final");
                        const done = [];
                        for (var i = 0; i < check.length; i++) {
                            const data = check[i].innerText.toLowerCase();
                          if (data.includes(".000₫") || data.includes(".000đ") || data.includes("mới ra mắt") 
                          	|| data.includes("tạm hết hàng") || data.includes("ngừng kinh doanh") 
                          	|| data.includes("hàng sắp về"))
                          {
                          	if(data.includes(".000₫") == true || data.includes(".000đ") == true )
                            {
                                const gia = /[0-9]/g;
                              	var first = data.search(gia);
                              	var last = data.search("₫");
                                done.push(data.slice(first,last));
                                break;
                            }else if(data.includes("mới ra mắt") == true || data.includes("tạm hết hàng") == true || data.includes("ngừng kinh doanh") == true 
                            	|| data.includes("hàng sắp về") == true){ 
                            	done.push(data)
                            	break;
                            }
                            else
                            {
                                done.push(data)
                                break;
                            }
                           // 
                          }
                        }return done; // trả về mảng dữ liệu
                    }
                    function checkGiaKM(){
                    	const checkppricesectionnk = document.querySelector(".productInfo_col-2")
                    	const checkppricesectiondmx = document.querySelector(".boxshock")
                    	if(checkppricesectionnk !== null || checkppricesectiondmx !== null){
                    		const check = document.querySelectorAll(".productInfo_col-2 .nk-shock-price, .boxshock .shockbuttonbox b");
	                        const done = [];
	                        for (var i = 0; i < check.length; i++) {
	                            const data = check[i].innerText.toLowerCase();
	                          if (data.includes(".000đ") || data.includes(".000₫"))
	                          {
	                                const gia = /[0-9]/g;
	                              	var first = data.search(gia);
	                              	var last = data.search("đ");
	                                done.push(data.slice(first,last));
	                                break;
	                          }
	                        }return done; // trả về mảng dữ liệu
                    	}
                    }
                    const price = checkGia();
                    const proprice = checkGiaKM();
                    const name = document.querySelector(checknamedmx !== null ? "h1" : checknamenk ? "h1.product_info_name" : ".khongcoclass")
                    

                    let data = {
                        Name: name ? name.innerText : "Không!",
                        Price: price ? price : "Không!",
                        Price2: proprice ? proprice : "Không!",
                    }
                   
                   
                    function checkTrutien()
                    {   
                        const check = document.querySelectorAll("span.promo, .shockbuttonbox, span.promo_text, .product-list-field p span a, .cm-picker-product-options label.labelForInput");
                        const done = [];
                        for (var i = 0; i < check.length; i++) {
                            const data = check[i].innerText.toLowerCase();
                          if (data.includes("khi thanh toán bằng thẻ tín dụng") 
                            || data.includes("hoàn tiền ngay") || data.includes("mua ngay") 
                            || data.includes("phiếu mua hàng") || data.includes("nkare") 
                            || data.includes("nhận phiếu mua hàng") || data.includes("thùng bia") 
                          	|| data.includes("online thêm quà") || data.includes("giảm ngay") 
                            || data.includes("vì cộng đồng vui khỏe") || data.includes("nkbest") 
                            || data.includes("giảm thêm") || data.includes("mua online") 
                            || data.includes("tháng tiền điện") || data.includes("giảm giá")
                            || data.includes("đã trừ vào giá"))
                          {
                          	if(data.includes("mua ngay") == true){
                          		done.push(data)
                          	}else if(data.includes("khi thanh toán bằng thẻ tín dụng") == true){
                          		done.push(null);
                          	}else if(data.includes("đã trừ vào giá") == true 
                          		|| data.includes("tháng tiền điện") == true 
                                || data.includes("giảm giá") ){
                                    var first = data.search("g");
                                    var last = data.search("giá");
                                    var rs = data.slice(0,60);
                                    done.push(rs.replace(/[:] /g,''));
                            }else if(data.includes("giảm giá") == true){
                          			var first = data.search("g");
	                              	var last = data.search("đ");
	                              	var rs = data.slice(first,last);
	                                done.push(rs.replace(/[:] /g,''));
                          	}else if(data.includes("tháng tiền điện") == true){
                                    // done.push(data.slice(1,41));
                                    var last = data.search("0đ");
                                    var last1 = last + 1
                                    var rs = data.slice(1,last1);
                                    done.push(rs.replace(/[:] /g,''));
                                    
                            }
                          	else{
                          		done.push(data.replace("\n",""));
                          	}
                          }
                        }return done; // trả về mảng dữ liệu
                    };

                     
                    const laygift = checkTrutien();
                    const gift = laygift[0];
                    const gift2 = laygift[1];
                    const gift3 = laygift[2];
                    const gift4 = laygift[3]; 
                    const tonggiatri = document.querySelector(checktonggiatri !== null && checktonggiatri1 === null ? ".area_promotion strong b" : checktonggiatri === null && checktonggiatri1 !== null ? ".khongcoclass" : ".khongcoclass")

                    return {
                        ...data,
                        Gift: gift ? gift : "0",
                        Gift2: gift2 ? gift2: "0",
                        Gift3: gift3 ? gift3 : "0",
                        Gift4: gift4 ? gift4 : "0",
                        Tonggiatri: tonggiatri ? tonggiatri.innerText : "0",
                    }
                }

                return {
                    Name: "Không Link!",
                    Price: "Không Link!",
                    Price2: "Không Link!",
                    Gift: "Không Link!",
                    Gift2: "Không Link!",
                    Gift3: "Không Link!",
                    Gift4: "Không Link!",
                    Tonggiatri: "Không Link!",
                };

            })
            if (info) {
                arrInfo.push(info)
                //res.writeHeader(200 , {"Content-Type" : "text/html; charset=utf-8"});
                res.write("<p>" + doithu.STT + " " + " => Xong!</p>");
            }
        } catch (err) {
            res.write("<p>Lỗi => " + doithu.STT +"</p>")
        }
    }
    var tenlink = exportToExcelDoiThu(arrInfo)
	exportToExcelDoiThu(arrInfo);
	res.write("<a href='/'>Back</a></br>");
	res.end("<a href='/"+tenlink+"'>Download File</a>");
	})();
	}
});
app.listen(process.env.PORT || 3000);