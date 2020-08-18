Bước 1: cài và triển khai app qua heroku
Bước 2
Để chạy puppeteer cần cài 
Cần phải cài build pack nodejs 

Sau đó cài tiếp build pack của puppeteer
$ heroku buildpacks:clear
$ heroku buildpacks:add --index 1 https://github.com/jontewks/puppeteer-heroku-buildpack
$ heroku buildpacks:add --index 1 heroku/nodejs

const browser = await puppeteer.launch({
  'args' : [
    '--no-sandbox',
    '--disable-setuid-sandbox'
  ]
});

$ git add .
$ git commit -am "cau commit"
$ git push heroku master

Để deploy source lên server dùng câu lệch sau 
$ git add . ( có thể bỏ qua câu lệnh git add . nếu đã deploy trước ")
$ git commit -am "cau commit" ( câu lệnh này dùng để check lại những gì đã update trong source )  
$ git push heroku master (dùng để deploy source)