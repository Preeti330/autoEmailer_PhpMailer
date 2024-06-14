# autoEmailer_PhpMailer
in this example create an excel sheet for list of stocks and share it to registred emails.and run cron job to scedule this email every day. 


#  steps
<ul>
  <li>
    Here first fetch all the records form inevtory table as total_qty,used_qty,remainings.
  </li>
  <li>
    After this genearte excel sheet using Spreedsheet . 
    to generate the excel sheet, and keep it dynamic header creaters . use ASCI Char concepets to auto fill headers and data.
  </li>
  <li> Install PHPMailer , via composer and set host,username,password and also set emails recivers ids </li>
</ul>
