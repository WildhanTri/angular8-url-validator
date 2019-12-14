import { Component, OnInit } from '@angular/core';
import * as XLSX from 'xlsx';
import readXlsxFile from 'read-excel-file'

@Component({
  selector: 'app-home',
  templateUrl: './home.component.html',
  styleUrls: ['./home.component.scss']
})
export class HomeComponent implements OnInit {

  urls = ""
  urlResults: UrlResult[] = []
  fileName = ""

  fileResults = []

  constructor() { }

  ngOnInit() {
  }

  checkUrls() {

    this.urlResults = []
    var manyUrls = this.urls.split(";");

    for (let url of manyUrls) {
      this.checkUrl(url)
    }
  }

  checkUrl(url) {
    //CLEAN URL IF HAVE WHITESPACE
    url = url.replace(/\s/g, "")

    var x = new XMLHttpRequest();
    x.timeout = 1000;
    x.open('GET', url);
    x.setRequestHeader("Access-Control-Allow-Headers", "x-requested-with, x-requested-by")
    x.onreadystatechange = (e: any) => {
      var http: XMLHttpRequest = e.currentTarget;
      if (http.readyState == 4) {
        console.log(http)
        if (http.status == 200) {
          var urlResult = new UrlResult();
          urlResult.url = url;
          urlResult.status = "OK";
          this.urlResults.push(urlResult)
        } else {
          var urlResult = new UrlResult();
          urlResult.url = url;
          urlResult.status = "Not OK";
          this.urlResults.push(urlResult)
        }
      }
    }
    x.send();
  }

  onFileSelected(event) {
    this.fileName = event.target.files[0].name
    readXlsxFile(event.target.files[0]).then((rows) => {
      console.log(rows)
      let index = 0;
      let header = []
      let results = []
      for (let r1 of rows) {
        if (index > 0) {
          let map = {}
          let index2 = 0
          for (let h of header) {
            map[h] = r1[index2]
            index2++;
          }
          results.push(map)
        } else {
          for (let r2 of r1) {
            header.push(r2)
          }
        }
        index++;
      }

      this.fileResults = []
      console.log(results)
      for (let r of results) {
        //CLEAN URL IF HAVE WHITESPACE

        if (r["url"] != null) {
          let url = r["url"].replace(/\s/g, "")

          var x = new XMLHttpRequest();
          x.timeout = 3000;
          x.open('GET', url);
          x.setRequestHeader("Access-Control-Allow-Headers", "x-requested-with, x-requested-by")
          x.onreadystatechange = (e: any) => {
            var http: XMLHttpRequest = e.currentTarget;
            if (http.readyState == 4) {
              if (http.status == 200) {
                // var urlResult = new UrlResult();
                // urlResult.url = url;
                // urlResult.status = "OK";
                // this.urlResults.push(urlResult)
                r["status_url"] = "OK"
                this.fileResults.push(r)
              } else {
                // var urlResult = new UrlResult();
                // urlResult.url = url;
                // urlResult.status = "Not OK";
                // this.urlResults.push(urlResult)
                r["status_url"] = "Not OK"
                this.fileResults.push(r)
              }
              if (this.fileResults.length == results.length) {
                this.reportToExcel(this.fileResults)
              }

            }
          }
          x.send();
        }

      }
    })
  }

  reportToExcel(json) {
    const workBook = XLSX.utils.book_new(); // create a new blank book
    const workSheet = XLSX.utils.json_to_sheet(json);
    var time = new Date()
    var filename = "Url-validator-report.xlsx"
    XLSX.utils.book_append_sheet(workBook, workSheet, 'Urls'); // add the worksheet to the book
    XLSX.writeFile(workBook, filename); // initiate a file download in browser
  }

}

class UrlResult {
  url: string;
  status: string;
}
