//#1
function addFooterImage()
{
  var doc = DocumentApp.getActiveDocument()
  //var body = doc.getBody();
  var footer = doc.getFooter()
  var image = UrlFetchApp.fetch('https://scontent.frwn2-1.fna.fbcdn.net/v/t39.30808-6/369604844_122097518594025492_955168146820002153_n.jpg?_nc_cat=101&ccb=1-7&_nc_sid=5f2048&_nc_ohc=QYds6jDEO4MAX8iELPR&_nc_ht=scontent.frwn2-1.fna&oh=00_AfCAfU34ql05R29I60oqk0EjID4lbKtmr68S93f8yvicpw&oe=65517AA2', {muteHttpExceptions: true,}).getBlob();
  footer.clear()
  var footerImage = footer.appendImage(image)
  footerImage.setHeight(50)
  footerImage.setWidth(50)
  footerImage.getParent().setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
}

//#2
function templateChange()
{
  const student = 
  [
    {
      surname: "Єніч",
      name: "Дмитро",
      patronomic: "Сергійович",
      group: 'ІПЗ-1',
      course: 'III',
      department: 'Харчових технологій'
    }
  ]

  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/1MzAMw8RlbVfwB_ttTqX_PEiw56bdcOwNdgHO-tryNhI/edit')
  var body = doc.getBody()
  var date = new Date();

  body.replaceText('<ПІБ>', student[0].surname + " " + student[0].name + " " + student[0].patronomic)
  body.replaceText('<курс>', student[0].course)
  body.replaceText('<факультет>', student[0].department)
  body.replaceText('<група>', student[0].course)
  body.replaceText('<дата>', date.getDate() + '.' + date.getMonth() + '.' + date.getFullYear())
}

//#3
function changeStudent(name)
{
  var doc = DocumentApp.openByUrl('https://docs.google.com/document/d/13-rkATVw64cOm6eflwKhq7PBDwPV8ysnc5FZJkY0EPA/edit')
  var allTextFromDoc = doc.getBody().getText()
  var body = doc.getBody()
  var template = /ІПЗ-1(.*?)Посилання/s
  var findText = allTextFromDoc.match(template)
  body.replaceText(findText[1].toString().trim(), name)
}

//#4
function linkToHyperlink() 
{
  var doc = DocumentApp.openByUrl("https://docs.google.com/document/d/1RNMuiEoscgUruUBfifyYeA4UomlovtJw5nLat6-yjSs/edit")
  var body = doc.getBody()
  var paragraphs = body.getParagraphs()
  var allTextFromDoc = doc.getBody().getText()
  var templateUrl = /https?:\/\/[^\s]+/g
  var allUrl = allTextFromDoc.match(templateUrl)
  console.log(allUrl)
  
  if(allUrl==null)
    return

  for(var i = 0; i<allUrl.length;i++)
  {
    var response = UrlFetchApp.fetch(allUrl[i], {muteHttpExceptions: true})
    var html = response.getContentText()
    var titleMatch = /<title>(.*?)<\/title>/i.exec(html)
    var title = titleMatch ? titleMatch[1] : url
    body.replaceText(allUrl[i], title)
  }
}

function main()
{
  changeStudent("Ісаєнко Віталій")
}
