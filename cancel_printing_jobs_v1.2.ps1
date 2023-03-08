
try {
  # Getting date to add it to email's subject
  $date = Get-Date -Format "yyyy-MM-dd (HH:mm)"

  # Preparing information's header
  $log_file_header = ""
  $log_file_header += "------------------------------------------------------------`r`n`r`n"
  $log_file_header += "   This is the report corresponding to $date`r`n`r`n"
  $log_file_header += "------------------------------------------------------------`r`n`r`n"
  $log_file_header += "`r`n`r`n"
  $log_file_header += "`r`n`r`n"
  $log_file_header += " >> Cancelling all jobs in the following printers: `r`n`r`n"

  # Writing header of the information to the file
  Set-Content -Path .\cancelled_printers.txt -Value $log_file_header -Encoding ASCII

  # Getting printers with Error State in:
  # 1 (Other error)
  # 6 (No toner)
  # 8 (Jammed)
  # 9 (Offline)
  $printers_with_errors = Get-WmiObject Win32_Printer -Filter "DetectedErrorState = 1 OR DetectedErrorState = 6 OR DetectedErrorState = 8 OR DetectedErrorState = 9 " 

  $rowsOfTable = ''

  $flag_send_email = $false

  # For every printer in the errored list, 
  # If there are print jobs for this printer
  # Append the printer name in the 'cancelled_printers.txt' file
  foreach ($printer in $printers_with_errors) {

    $printerName = $printer.Name

    $erroredPrinter = Get-WmiObject Win32_PerfFormattedData_Spooler_PrintQueue -Filter "Name LIKE '$printerName'"
    $printerWithErrorAndJobs = $erroredPrinter | Select-Object Name, Jobs, Availability | Where-Object { $_.jobs -gt 0 }
    
    Write-Output $printerWithErrorAndJobs

    # A flag to know when to send email...
    #$flag_send_email = $false

    if ($erroredPrinter.Jobs -gt 0) {
        
      $rowsOfTable += '<tr><td class="tg-phtq">' + $erroredPrinter.Name.toUpper() + '</td>'
      $rowsOfTable += '<td class="tg-phtq">' + $erroredPrinter.Jobs + '</td></tr>'
        
      $outputToTxtFile = $erroredPrinter.Name.toUpper() + " had " + $erroredPrinter.Jobs + " print job(s) that were cancelled...`n"
      Out-File -InputObject $outputToTxtFile -FilePath .\cancelled_printers.txt -Append -Encoding ASCII
        
      $printer.CancelAllJobs()

      $flag_send_email = $true
    }
  }

  Write-Output "Flag value: "$flag_send_email

  $EmailSmtpServer = 'mail.server.com'
  $EmailFrom = 'Printers Watcher <no-reply@server.com>'
  $EmailTo = 'IT guy <itguy@server.com>'
  $EmailSubject = "$date - Cancelled printing jobs daily report"
  $EmailSubject = "Errored print jobs report - $date" 


  # If the flag for sending email is setted to true, send the email
  if ($flag_send_email -eq $true) {
    $EmailBody = '<style type="text/css">
    p,h1 {
        font-family: "Segoe UI";
    }
    .tg  {border-collapse:collapse;border-color:#9ABAD9;border-spacing:0;margin:0px auto;}
    .tg td{background-color:#EBF5FF;border-color:#9ABAD9;border-style:solid;border-width:1px;color:#444;
      font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;word-break:normal;}
    .tg th{background-color:#409cff;border-color:#9ABAD9;border-style:solid;border-width:1px;color:#fff;
      font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 5px;word-break:normal;}
    .tg .tg-phtq{background-color:#D2E4FC;border-color:inherit;text-align:left;vertical-align:top}
    .tg .tg-9abe{border-color:inherit;font-size:16px;font-weight:bold;position:-webkit-sticky;position:sticky;text-align:left;top:-1px;
      vertical-align:top;will-change:transform}
    .tg-sort-header::-moz-selection{background:0 0}
    .tg-sort-header::selection{background:0 0}.tg-sort-header{cursor:pointer}
    .tg-sort-header:after{content:'';float:right;margin-top:7px;border-width:0 5px 5px;border-style:solid;
      border-color:#404040 transparent;visibility:hidden}
    .tg-sort-header:hover:after{visibility:visible}
    .tg-sort-asc:after,.tg-sort-asc:hover:after,.tg-sort-desc:after{visibility:visible;opacity:.4}
    .tg-sort-desc:after{border-bottom:none;border-width:5px 5px 0}@media screen and (max-width: 767px) {.tg {width: auto !important;}.tg col {width: auto !important;}.tg-wrap {overflow-x: auto;-webkit-overflow-scrolling: touch;margin: auto 0px;}}</style>
    <h1 class="tg-9abe">Impresoras con errores</h1>
    <p>
        De las impresoras de <strong>SERVER-PSX</strong> en estado de error, fuera de l&iacute;nea o atascadas, se han cancelado todos los trabajos de impresi&oacute;n almacenados en las impresoras mostradas en la tabla.
    </p>
    <div class="tg-wrap"><table id="tg-wBO1i" class="tg">
    <thead>
      <tr>
        <th class="tg-9abe">Nombre de Impresora</th>
        <th class="tg-9abe">Cantidad de trabajos cancelados</th>
      </tr>
    </thead>
    <tbody>
        ' + $rowsOfTable + '
    </tbody>
    </table></div>
    <script charset="utf-8">var TGSort=window.TGSort||function(n){"use strict";function r(n){return n?n.length:0}function t(n,t,e,o=0){for(e=r(n);o<e;++o)t(n[o],o)}function e(n){return n.split("").reverse().join("")}function o(n){var e=n[0];return t(n,function(n){for(;!n.startsWith(e);)e=e.substring(0,r(e)-1)}),r(e)}function u(n,r,e=[]){return t(n,function(n){r(n)&&e.push(n)}),e}var a=parseFloat;function i(n,r){return function(t){var e="";return t.replace(n,function(n,t,o){return e=t.replace(r,"")+"."+(o||"").substring(1)}),a(e)}}var s=i(/^(?:\s*)([+-]?(?:\d+)(?:,\d{3})*)(\.\d*)?$/g,/,/g),c=i(/^(?:\s*)([+-]?(?:\d+)(?:\.\d{3})*)(,\d*)?$/g,/\./g);function f(n){var t=a(n);return!isNaN(t)&&r(""+t)+1>=r(n)?t:NaN}function d(n){var e=[],o=n;return t([f,s,c],function(u){var a=[],i=[];t(n,function(n,r){r=u(n),a.push(r),r||i.push(n)}),r(i)<r(o)&&(o=i,e=a)}),r(u(o,function(n){return n==o[0]}))==r(o)?e:[]}function v(n){if("TABLE"==n.nodeName){for(var a=function(r){var e,o,u=[],a=[];return function n(r,e){e(r),t(r.childNodes,function(r){n(r,e)})}(n,function(n){"TR"==(o=n.nodeName)?(e=[],u.push(e),a.push(n)):"TD"!=o&&"TH"!=o||e.push(n)}),[u,a]}(),i=a[0],s=a[1],c=r(i),f=c>1&&r(i[0])<r(i[1])?1:0,v=f+1,p=i[f],h=r(p),l=[],g=[],N=[],m=v;m<c;++m){for(var T=0;T<h;++T){r(g)<h&&g.push([]);var C=i[m][T],L=C.textContent||C.innerText||"";g[T].push(L.trim())}N.push(m-v)}t(p,function(n,t){l[t]=0;var a=n.classList;a.add("tg-sort-header"),n.addEventListener("click",function(){var n=l[t];!function(){for(var n=0;n<h;++n){var r=p[n].classList;r.remove("tg-sort-asc"),r.remove("tg-sort-desc"),l[n]=0}}(),(n=1==n?-1:+!n)&&a.add(n>0?"tg-sort-asc":"tg-sort-desc"),l[t]=n;var i,f=g[t],m=function(r,t){return n*f[r].localeCompare(f[t])||n*(r-t)},T=function(n){var t=d(n);if(!r(t)){var u=o(n),a=o(n.map(e));t=d(n.map(function(n){return n.substring(u,r(n)-a)}))}return t}(f);(r(T)||r(T=r(u(i=f.map(Date.parse),isNaN))?[]:i))&&(m=function(r,t){var e=T[r],o=T[t],u=isNaN(e),a=isNaN(o);return u&&a?0:u?-n:a?n:e>o?n:e<o?-n:n*(r-t)});var C,L=N.slice();L.sort(m);for(var E=v;E<c;++E)(C=s[E].parentNode).removeChild(s[E]);for(E=v;E<c;++E)C.appendChild(s[v+L[E-v]])})})}}n.addEventListener("DOMContentLoaded",function(){for(var t=n.getElementsByClassName("tg"),e=0;e<r(t);++e)try{v(t[e])}catch(n){}})}(document)</script>'
  
    Send-MailMessage -From $EmailFrom -To $EmailTo -Subject $EmailSubject -SmtpServer $EmailSmtpServer -Body $EmailBody -BodyAsHtml
  }
  else {
    # Do nothing until now
  }
}
catch {
  Write-Host "`n`t=====================================================================================================================`n" -ForegroundColor Red
  Write-Host "`t>>> Error: $($_)`n" -ForegroundColor Red
  Write-Host "`t=====================================================================================================================`n" -ForegroundColor Red
}
