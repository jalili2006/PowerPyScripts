# Check if the module is installed
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    # Install the module automatically without any prompt
    Install-Module -Name ImportExcel -Force -Confirm:$false
}

# Import the Excel module
Import-Module ImportExcel

$file1 = 'C:\Users\jalil\Documents\1.xlsx' # source's fullpath 

$xl=New-Object -ComObject Excel.Application
$xl.Visible=$true
$wb=$xl.WorkBooks.Open('C:\Users\jalil\Documents\1.xlsx') #Target
$ws=$wb.WorkSheets.item(1)

$wb2 = $xl.workbooks.open($file1, $null, $true) # open source, readonly
$sheetToCopy = $wb2.sheets.item(1) # source sheet to copy 

#-------------------------------------------------------------------------------------------
#Check Source row less than 100
Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$xl.Visible = $false

$ws2 = $wb2.Worksheets.Item(1)

$objRange = $ws2.UsedRange
$a1 = $objRange.SpecialCells(11).row
$b1 = $objRange.SpecialCells(11).column

if($a1 -ge 101){

#write-host "It is too big, Please reduce the number of rows less than 100"

write-host "Lastrow:", $a1, " Last Column:" $b1 

[System.Windows.MessageBox]::Show('Source is too big...!, The total number of rows is: '+$a1+' Please reduce the number of rows less than 100','Error','OK','Error')

$xl.Quit($false)
#$wb2.close($false) #Source
#$wb.close($false) #destination 
#$xl.quit()
# [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)

Break

}


#--------------------------------------------------------------------------------------------
$sor1 = $sheetToCopy.Cells.Item(1,1)                     #Source
$sor2 = $sheetToCopy.Cells.Item(1,2)
$sor3 = $sheetToCopy.Cells.Item(1,3)
$sor4 = $sheetToCopy.Cells.Item(1,4)
$sor5 = $sheetToCopy.Cells.Item(1,5)
$sor6 = $sheetToCopy.Cells.Item(1,6)
$sor7 = $sheetToCopy.Cells.Item(1,7)           
$sorlast = $sheetToCopy.Cells.Item(1,8)

$sor21 = $sheetToCopy.Cells.Item(2,1)                     #Source for second row
$sor22 = $sheetToCopy.Cells.Item(2,2)
$sor23 = $sheetToCopy.Cells.Item(2,3)
$sor24 = $sheetToCopy.Cells.Item(2,4)
$sor25 = $sheetToCopy.Cells.Item(2,5)
$sor26 = $sheetToCopy.Cells.Item(2,6)
$sor27 = $sheetToCopy.Cells.Item(2,7)           
$sor2last = $sheetToCopy.Cells.Item(2,8)

$sor31 = $sheetToCopy.Cells.Item(3,1)                     #Source for 3th row
$sor32 = $sheetToCopy.Cells.Item(3,2)
$sor33 = $sheetToCopy.Cells.Item(3,3)
$sor34 = $sheetToCopy.Cells.Item(3,4)
$sor35 = $sheetToCopy.Cells.Item(3,5)
$sor36 = $sheetToCopy.Cells.Item(3,6)
$sor37 = $sheetToCopy.Cells.Item(3,7)           
$sor3last = $sheetToCopy.Cells.Item(3,8)

$sor41 = $sheetToCopy.Cells.Item(4,1)                     #Source for 4th row
$sor42 = $sheetToCopy.Cells.Item(4,2)
$sor43 = $sheetToCopy.Cells.Item(4,3)
$sor44 = $sheetToCopy.Cells.Item(4,4)
$sor45 = $sheetToCopy.Cells.Item(4,5)
$sor46 = $sheetToCopy.Cells.Item(4,6)
$sor47 = $sheetToCopy.Cells.Item(4,7)           
$sor4last = $sheetToCopy.Cells.Item(4,8)

$sor51 = $sheetToCopy.Cells.Item(5,1)                     #Source for 5th row
$sor52 = $sheetToCopy.Cells.Item(5,2)
$sor53 = $sheetToCopy.Cells.Item(5,3)
$sor54 = $sheetToCopy.Cells.Item(5,4)
$sor55 = $sheetToCopy.Cells.Item(5,5)
$sor56 = $sheetToCopy.Cells.Item(5,6)
$sor57 = $sheetToCopy.Cells.Item(5,7)           
$sor5last = $sheetToCopy.Cells.Item(5,8)

$sor61 = $sheetToCopy.Cells.Item(6,1)                     #Source for 6th row
$sor62 = $sheetToCopy.Cells.Item(6,2)
$sor63 = $sheetToCopy.Cells.Item(6,3)
$sor64 = $sheetToCopy.Cells.Item(6,4)
$sor65 = $sheetToCopy.Cells.Item(6,5)
$sor66 = $sheetToCopy.Cells.Item(6,6)
$sor67 = $sheetToCopy.Cells.Item(6,7)           
$sor6last = $sheetToCopy.Cells.Item(6,8)

$sor71 = $sheetToCopy.Cells.Item(7,1)                     #Source for 7th row
$sor72 = $sheetToCopy.Cells.Item(7,2)
$sor73 = $sheetToCopy.Cells.Item(7,3)
$sor74 = $sheetToCopy.Cells.Item(7,4)
$sor75 = $sheetToCopy.Cells.Item(7,5)
$sor76 = $sheetToCopy.Cells.Item(7,6)
$sor77 = $sheetToCopy.Cells.Item(7,7)           
$sor7last = $sheetToCopy.Cells.Item(7,8)

$sor81 = $sheetToCopy.Cells.Item(8,1)                     #Source for 8th row
$sor82 = $sheetToCopy.Cells.Item(8,2)
$sor83 = $sheetToCopy.Cells.Item(8,3)
$sor84 = $sheetToCopy.Cells.Item(8,4)
$sor85 = $sheetToCopy.Cells.Item(8,5)
$sor86 = $sheetToCopy.Cells.Item(8,6)
$sor87 = $sheetToCopy.Cells.Item(8,7)           
$sor8last = $sheetToCopy.Cells.Item(8,8)

$sor91 = $sheetToCopy.Cells.Item(9,1)                     #Source for 9th row
$sor92 = $sheetToCopy.Cells.Item(9,2)
$sor93 = $sheetToCopy.Cells.Item(9,3)
$sor94 = $sheetToCopy.Cells.Item(9,4)
$sor95 = $sheetToCopy.Cells.Item(9,5)
$sor96 = $sheetToCopy.Cells.Item(9,6)
$sor97 = $sheetToCopy.Cells.Item(9,7)           
$sor9last = $sheetToCopy.Cells.Item(9,8)

$sor101 = $sheetToCopy.Cells.Item(10,1)                     #Source for 10th row
$sor102 = $sheetToCopy.Cells.Item(10,2)
$sor103 = $sheetToCopy.Cells.Item(10,3)
$sor104 = $sheetToCopy.Cells.Item(10,4)
$sor105 = $sheetToCopy.Cells.Item(10,5)
$sor106 = $sheetToCopy.Cells.Item(10,6)
$sor107 = $sheetToCopy.Cells.Item(10,7)           
$sor10last = $sheetToCopy.Cells.Item(10,8)

$sor111 = $sheetToCopy.Cells.Item(11,1)                     #Source for 11th row
$sor112 = $sheetToCopy.Cells.Item(11,2)
$sor113 = $sheetToCopy.Cells.Item(11,3)
$sor114 = $sheetToCopy.Cells.Item(11,4)
$sor115 = $sheetToCopy.Cells.Item(11,5)
$sor116 = $sheetToCopy.Cells.Item(11,6)
$sor117 = $sheetToCopy.Cells.Item(11,7)           
$sor11last = $sheetToCopy.Cells.Item(11,8)

$sor121 = $sheetToCopy.Cells.Item(12,1)                     #Source for 12th row
$sor122 = $sheetToCopy.Cells.Item(12,2)
$sor123 = $sheetToCopy.Cells.Item(12,3)
$sor124 = $sheetToCopy.Cells.Item(12,4)
$sor125 = $sheetToCopy.Cells.Item(12,5)
$sor126 = $sheetToCopy.Cells.Item(12,6)
$sor127 = $sheetToCopy.Cells.Item(12,7)           
$sor12last = $sheetToCopy.Cells.Item(12,8)

$sor131 = $sheetToCopy.Cells.Item(13,1)                     #Source for 13th row
$sor132 = $sheetToCopy.Cells.Item(13,2)
$sor133 = $sheetToCopy.Cells.Item(13,3)
$sor134 = $sheetToCopy.Cells.Item(13,4)
$sor135 = $sheetToCopy.Cells.Item(13,5)
$sor136 = $sheetToCopy.Cells.Item(13,6)
$sor137 = $sheetToCopy.Cells.Item(13,7)           
$sor13last = $sheetToCopy.Cells.Item(13,8)

$sor141 = $sheetToCopy.Cells.Item(14,1)                     #Source for 14th row
$sor142 = $sheetToCopy.Cells.Item(14,2)
$sor143 = $sheetToCopy.Cells.Item(14,3)
$sor144 = $sheetToCopy.Cells.Item(14,4)
$sor145 = $sheetToCopy.Cells.Item(14,5)
$sor146 = $sheetToCopy.Cells.Item(14,6)
$sor147 = $sheetToCopy.Cells.Item(14,7)           
$sor14last = $sheetToCopy.Cells.Item(14,8)

$sor151 = $sheetToCopy.Cells.Item(15,1)                     #Source for 15th row
$sor152 = $sheetToCopy.Cells.Item(15,2)
$sor153 = $sheetToCopy.Cells.Item(15,3)
$sor154 = $sheetToCopy.Cells.Item(15,4)
$sor155 = $sheetToCopy.Cells.Item(15,5)
$sor156 = $sheetToCopy.Cells.Item(15,6)
$sor157 = $sheetToCopy.Cells.Item(15,7)           
$sor15last = $sheetToCopy.Cells.Item(15,8)

$sor161 = $sheetToCopy.Cells.Item(16,1)                     #Source for 16th row
$sor162 = $sheetToCopy.Cells.Item(16,2)
$sor163 = $sheetToCopy.Cells.Item(16,3)
$sor164 = $sheetToCopy.Cells.Item(16,4)
$sor165 = $sheetToCopy.Cells.Item(16,5)
$sor166 = $sheetToCopy.Cells.Item(16,6)
$sor167 = $sheetToCopy.Cells.Item(16,7)           
$sor16last = $sheetToCopy.Cells.Item(16,8)

$sor171 = $sheetToCopy.Cells.Item(17,1)                     #Source for 17th row
$sor172 = $sheetToCopy.Cells.Item(17,2)
$sor173 = $sheetToCopy.Cells.Item(17,3)
$sor174 = $sheetToCopy.Cells.Item(17,4)
$sor175 = $sheetToCopy.Cells.Item(17,5)
$sor176 = $sheetToCopy.Cells.Item(17,6)
$sor177 = $sheetToCopy.Cells.Item(17,7)           
$sor17last = $sheetToCopy.Cells.Item(17,8)

$sor181 = $sheetToCopy.Cells.Item(18,1)                     #Source for 18th row
$sor182 = $sheetToCopy.Cells.Item(18,2)
$sor183 = $sheetToCopy.Cells.Item(18,3)
$sor184 = $sheetToCopy.Cells.Item(18,4)
$sor185 = $sheetToCopy.Cells.Item(18,5)
$sor186 = $sheetToCopy.Cells.Item(18,6)
$sor187 = $sheetToCopy.Cells.Item(18,7)           
$sor18last = $sheetToCopy.Cells.Item(18,8)

$sor191 = $sheetToCopy.Cells.Item(19,1)                     #Source for 19th row
$sor192 = $sheetToCopy.Cells.Item(19,2)
$sor193 = $sheetToCopy.Cells.Item(19,3)
$sor194 = $sheetToCopy.Cells.Item(19,4)
$sor195 = $sheetToCopy.Cells.Item(19,5)
$sor196 = $sheetToCopy.Cells.Item(19,6)
$sor197 = $sheetToCopy.Cells.Item(19,7)           
$sor19last = $sheetToCopy.Cells.Item(19,8)

$sor201 = $sheetToCopy.Cells.Item(20,1)                     #Source for 20th row
$sor202 = $sheetToCopy.Cells.Item(20,2)
$sor203 = $sheetToCopy.Cells.Item(20,3)
$sor204 = $sheetToCopy.Cells.Item(20,4)
$sor205 = $sheetToCopy.Cells.Item(20,5)
$sor206 = $sheetToCopy.Cells.Item(20,6)
$sor207 = $sheetToCopy.Cells.Item(20,7)           
$sor20last = $sheetToCopy.Cells.Item(20,8)

$sor211 = $sheetToCopy.Cells.Item(21,1)                     #Source for 21th row
$sor212 = $sheetToCopy.Cells.Item(21,2)
$sor213 = $sheetToCopy.Cells.Item(21,3)
$sor214 = $sheetToCopy.Cells.Item(21,4)
$sor215 = $sheetToCopy.Cells.Item(21,5)
$sor216 = $sheetToCopy.Cells.Item(21,6)
$sor217 = $sheetToCopy.Cells.Item(21,7)           
$sor21last = $sheetToCopy.Cells.Item(21,8)

$sor221 = $sheetToCopy.Cells.Item(22,1)                     #Source for 22th row
$sor222 = $sheetToCopy.Cells.Item(22,2)
$sor223 = $sheetToCopy.Cells.Item(22,3)
$sor224 = $sheetToCopy.Cells.Item(22,4)
$sor225 = $sheetToCopy.Cells.Item(22,5)
$sor226 = $sheetToCopy.Cells.Item(22,6)
$sor227 = $sheetToCopy.Cells.Item(22,7)           
$sor22last = $sheetToCopy.Cells.Item(22,8)

$sor231 = $sheetToCopy.Cells.Item(23,1)                     #Source for 23th row
$sor232 = $sheetToCopy.Cells.Item(23,2)
$sor233 = $sheetToCopy.Cells.Item(23,3)
$sor234 = $sheetToCopy.Cells.Item(23,4)
$sor235 = $sheetToCopy.Cells.Item(23,5)
$sor236 = $sheetToCopy.Cells.Item(23,6)
$sor237 = $sheetToCopy.Cells.Item(23,7)           
$sor23last = $sheetToCopy.Cells.Item(23,8)

$sor241 = $sheetToCopy.Cells.Item(24,1)                     #Source for 24th row
$sor242 = $sheetToCopy.Cells.Item(24,2)
$sor243 = $sheetToCopy.Cells.Item(24,3)
$sor244 = $sheetToCopy.Cells.Item(24,4)
$sor245 = $sheetToCopy.Cells.Item(24,5)
$sor246 = $sheetToCopy.Cells.Item(24,6)
$sor247 = $sheetToCopy.Cells.Item(24,7)           
$sor24last = $sheetToCopy.Cells.Item(24,8)

$sor251 = $sheetToCopy.Cells.Item(25,1)                     #Source for 25th row
$sor252 = $sheetToCopy.Cells.Item(25,2)
$sor253 = $sheetToCopy.Cells.Item(25,3)
$sor254 = $sheetToCopy.Cells.Item(25,4)
$sor255 = $sheetToCopy.Cells.Item(25,5)
$sor256 = $sheetToCopy.Cells.Item(25,6)
$sor257 = $sheetToCopy.Cells.Item(25,7)           
$sor25last = $sheetToCopy.Cells.Item(25,8)

$sor261 = $sheetToCopy.Cells.Item(26,1)                     #Source for 26th row
$sor262 = $sheetToCopy.Cells.Item(26,2)
$sor263 = $sheetToCopy.Cells.Item(26,3)
$sor264 = $sheetToCopy.Cells.Item(26,4)
$sor265 = $sheetToCopy.Cells.Item(26,5)
$sor266 = $sheetToCopy.Cells.Item(26,6)
$sor267 = $sheetToCopy.Cells.Item(26,7)           
$sor26last = $sheetToCopy.Cells.Item(26,8)

$sor271 = $sheetToCopy.Cells.Item(27,1)                     #Source for 27th row
$sor272 = $sheetToCopy.Cells.Item(27,2)
$sor273 = $sheetToCopy.Cells.Item(27,3)
$sor274 = $sheetToCopy.Cells.Item(27,4)
$sor275 = $sheetToCopy.Cells.Item(27,5)
$sor276 = $sheetToCopy.Cells.Item(27,6)
$sor277 = $sheetToCopy.Cells.Item(27,7)           
$sor27last = $sheetToCopy.Cells.Item(27,8)

$sor281 = $sheetToCopy.Cells.Item(28,1)                     #Source for 28th row
$sor282 = $sheetToCopy.Cells.Item(28,2)
$sor283 = $sheetToCopy.Cells.Item(28,3)
$sor284 = $sheetToCopy.Cells.Item(28,4)
$sor285 = $sheetToCopy.Cells.Item(28,5)
$sor286 = $sheetToCopy.Cells.Item(28,6)
$sor287 = $sheetToCopy.Cells.Item(28,7)           
$sor28last = $sheetToCopy.Cells.Item(28,8)

$sor291 = $sheetToCopy.Cells.Item(29,1)                     #Source for 29th row
$sor292 = $sheetToCopy.Cells.Item(29,2)
$sor293 = $sheetToCopy.Cells.Item(29,3)
$sor294 = $sheetToCopy.Cells.Item(29,4)
$sor295 = $sheetToCopy.Cells.Item(29,5)
$sor296 = $sheetToCopy.Cells.Item(29,6)
$sor297 = $sheetToCopy.Cells.Item(29,7)           
$sor29last = $sheetToCopy.Cells.Item(29,8)

$sor301 = $sheetToCopy.Cells.Item(30,1)                     #Source for 30th row
$sor302 = $sheetToCopy.Cells.Item(30,2)
$sor303 = $sheetToCopy.Cells.Item(30,3)
$sor304 = $sheetToCopy.Cells.Item(30,4)
$sor305 = $sheetToCopy.Cells.Item(30,5)
$sor306 = $sheetToCopy.Cells.Item(30,6)
$sor307 = $sheetToCopy.Cells.Item(30,7)           
$sor30last = $sheetToCopy.Cells.Item(30,8)

$sor311 = $sheetToCopy.Cells.Item(31,1)                     #Source for 31th row
$sor312 = $sheetToCopy.Cells.Item(31,2)
$sor313 = $sheetToCopy.Cells.Item(31,3)
$sor314 = $sheetToCopy.Cells.Item(31,4)
$sor315 = $sheetToCopy.Cells.Item(31,5)
$sor316 = $sheetToCopy.Cells.Item(31,6)
$sor317 = $sheetToCopy.Cells.Item(31,7)           
$sor31last = $sheetToCopy.Cells.Item(31,8)

$sor321 = $sheetToCopy.Cells.Item(32,1)                     #Source for 32th row
$sor322 = $sheetToCopy.Cells.Item(32,2)
$sor323 = $sheetToCopy.Cells.Item(32,3)
$sor324 = $sheetToCopy.Cells.Item(32,4)
$sor325 = $sheetToCopy.Cells.Item(32,5)
$sor326 = $sheetToCopy.Cells.Item(32,6)
$sor327 = $sheetToCopy.Cells.Item(32,7)           
$sor32last = $sheetToCopy.Cells.Item(32,8)

$sor331 = $sheetToCopy.Cells.Item(33,1)                     #Source for 33th row
$sor332 = $sheetToCopy.Cells.Item(33,2)
$sor333 = $sheetToCopy.Cells.Item(33,3)
$sor334 = $sheetToCopy.Cells.Item(33,4)
$sor335 = $sheetToCopy.Cells.Item(33,5)
$sor336 = $sheetToCopy.Cells.Item(33,6)
$sor337 = $sheetToCopy.Cells.Item(33,7)           
$sor33last = $sheetToCopy.Cells.Item(33,8)

$sor341 = $sheetToCopy.Cells.Item(34,1)                     #Source for 34th row
$sor342 = $sheetToCopy.Cells.Item(34,2)
$sor343 = $sheetToCopy.Cells.Item(34,3)
$sor344 = $sheetToCopy.Cells.Item(34,4)
$sor345 = $sheetToCopy.Cells.Item(34,5)
$sor346 = $sheetToCopy.Cells.Item(34,6)
$sor347 = $sheetToCopy.Cells.Item(34,7)           
$sor34last = $sheetToCopy.Cells.Item(34,8)

$sor351 = $sheetToCopy.Cells.Item(35,1)                     #Source for 35th row
$sor352 = $sheetToCopy.Cells.Item(35,2)
$sor353 = $sheetToCopy.Cells.Item(35,3)
$sor354 = $sheetToCopy.Cells.Item(35,4)
$sor355 = $sheetToCopy.Cells.Item(35,5)
$sor356 = $sheetToCopy.Cells.Item(35,6)
$sor357 = $sheetToCopy.Cells.Item(35,7)           
$sor35last = $sheetToCopy.Cells.Item(35,8)

$sor361 = $sheetToCopy.Cells.Item(36,1)                     #Source for 36th row
$sor362 = $sheetToCopy.Cells.Item(36,2)
$sor363 = $sheetToCopy.Cells.Item(36,3)
$sor364 = $sheetToCopy.Cells.Item(36,4)
$sor365 = $sheetToCopy.Cells.Item(36,5)
$sor366 = $sheetToCopy.Cells.Item(36,6)
$sor367 = $sheetToCopy.Cells.Item(36,7)           
$sor36last = $sheetToCopy.Cells.Item(36,8)

$sor371 = $sheetToCopy.Cells.Item(37,1)                     #Source for 37th row
$sor372 = $sheetToCopy.Cells.Item(37,2)
$sor373 = $sheetToCopy.Cells.Item(37,3)
$sor374 = $sheetToCopy.Cells.Item(37,4)
$sor375 = $sheetToCopy.Cells.Item(37,5)
$sor376 = $sheetToCopy.Cells.Item(37,6)
$sor377 = $sheetToCopy.Cells.Item(37,7)           
$sor37last = $sheetToCopy.Cells.Item(37,8)

$sor381 = $sheetToCopy.Cells.Item(38,1)                     #Source for 38th row
$sor382 = $sheetToCopy.Cells.Item(38,2)
$sor383 = $sheetToCopy.Cells.Item(38,3)
$sor384 = $sheetToCopy.Cells.Item(38,4)
$sor385 = $sheetToCopy.Cells.Item(38,5)
$sor386 = $sheetToCopy.Cells.Item(38,6)
$sor387 = $sheetToCopy.Cells.Item(38,7)           
$sor38last = $sheetToCopy.Cells.Item(38,8)

$sor391 = $sheetToCopy.Cells.Item(39,1)                     #Source for 39th row
$sor392 = $sheetToCopy.Cells.Item(39,2)
$sor393 = $sheetToCopy.Cells.Item(39,3)
$sor394 = $sheetToCopy.Cells.Item(39,4)
$sor395 = $sheetToCopy.Cells.Item(39,5)
$sor396 = $sheetToCopy.Cells.Item(39,6)
$sor397 = $sheetToCopy.Cells.Item(39,7)           
$sor39last = $sheetToCopy.Cells.Item(39,8)

$sor401 = $sheetToCopy.Cells.Item(40,1)                     #Source for 40th row
$sor402 = $sheetToCopy.Cells.Item(40,2)
$sor403 = $sheetToCopy.Cells.Item(40,3)
$sor404 = $sheetToCopy.Cells.Item(40,4)
$sor405 = $sheetToCopy.Cells.Item(40,5)
$sor406 = $sheetToCopy.Cells.Item(40,6)
$sor407 = $sheetToCopy.Cells.Item(40,7)           
$sor40last = $sheetToCopy.Cells.Item(40,8)

$sor411 = $sheetToCopy.Cells.Item(41,1)                     #Source for 41th row
$sor412 = $sheetToCopy.Cells.Item(41,2)
$sor413 = $sheetToCopy.Cells.Item(41,3)
$sor414 = $sheetToCopy.Cells.Item(41,4)
$sor415 = $sheetToCopy.Cells.Item(41,5)
$sor416 = $sheetToCopy.Cells.Item(41,6)
$sor417 = $sheetToCopy.Cells.Item(41,7)           
$sor41last = $sheetToCopy.Cells.Item(41,8)

$sor421 = $sheetToCopy.Cells.Item(42,1)                     #Source for 42th row
$sor422 = $sheetToCopy.Cells.Item(42,2)
$sor423 = $sheetToCopy.Cells.Item(42,3)
$sor424 = $sheetToCopy.Cells.Item(42,4)
$sor425 = $sheetToCopy.Cells.Item(42,5)
$sor426 = $sheetToCopy.Cells.Item(42,6)
$sor427 = $sheetToCopy.Cells.Item(42,7)           
$sor42last = $sheetToCopy.Cells.Item(42,8)

$sor431 = $sheetToCopy.Cells.Item(43,1)                     #Source for 43th row
$sor432 = $sheetToCopy.Cells.Item(43,2)
$sor433 = $sheetToCopy.Cells.Item(43,3)
$sor434 = $sheetToCopy.Cells.Item(43,4)
$sor435 = $sheetToCopy.Cells.Item(43,5)
$sor436 = $sheetToCopy.Cells.Item(43,6)
$sor437 = $sheetToCopy.Cells.Item(43,7)           
$sor43last = $sheetToCopy.Cells.Item(43,8)

$sor441 = $sheetToCopy.Cells.Item(44,1)                     #Source for 44th row
$sor442 = $sheetToCopy.Cells.Item(44,2)
$sor443 = $sheetToCopy.Cells.Item(44,3)
$sor444 = $sheetToCopy.Cells.Item(44,4)
$sor445 = $sheetToCopy.Cells.Item(44,5)
$sor446 = $sheetToCopy.Cells.Item(44,6)
$sor447 = $sheetToCopy.Cells.Item(44,7)           
$sor44last = $sheetToCopy.Cells.Item(44,8)

$sor451 = $sheetToCopy.Cells.Item(45,1)                     #Source for 45th row
$sor452 = $sheetToCopy.Cells.Item(45,2)
$sor453 = $sheetToCopy.Cells.Item(45,3)
$sor454 = $sheetToCopy.Cells.Item(45,4)
$sor455 = $sheetToCopy.Cells.Item(45,5)
$sor456 = $sheetToCopy.Cells.Item(45,6)
$sor457 = $sheetToCopy.Cells.Item(45,7)           
$sor45last = $sheetToCopy.Cells.Item(45,8)

$sor461 = $sheetToCopy.Cells.Item(46,1)                     #Source for 46th row
$sor462 = $sheetToCopy.Cells.Item(46,2)
$sor463 = $sheetToCopy.Cells.Item(46,3)
$sor464 = $sheetToCopy.Cells.Item(46,4)
$sor465 = $sheetToCopy.Cells.Item(46,5)
$sor466 = $sheetToCopy.Cells.Item(46,6)
$sor467 = $sheetToCopy.Cells.Item(46,7)           
$sor46last = $sheetToCopy.Cells.Item(46,8)

$sor471 = $sheetToCopy.Cells.Item(47,1)                     #Source for 47th row
$sor472 = $sheetToCopy.Cells.Item(47,2)
$sor473 = $sheetToCopy.Cells.Item(47,3)
$sor474 = $sheetToCopy.Cells.Item(47,4)
$sor475 = $sheetToCopy.Cells.Item(47,5)
$sor476 = $sheetToCopy.Cells.Item(47,6)
$sor477 = $sheetToCopy.Cells.Item(47,7)           
$sor47last = $sheetToCopy.Cells.Item(47,8)

$sor481 = $sheetToCopy.Cells.Item(48,1)                     #Source for 48th row
$sor482 = $sheetToCopy.Cells.Item(48,2)
$sor483 = $sheetToCopy.Cells.Item(48,3)
$sor484 = $sheetToCopy.Cells.Item(48,4)
$sor485 = $sheetToCopy.Cells.Item(48,5)
$sor486 = $sheetToCopy.Cells.Item(48,6)
$sor487 = $sheetToCopy.Cells.Item(48,7)           
$sor48last = $sheetToCopy.Cells.Item(48,8)

$sor491 = $sheetToCopy.Cells.Item(49,1)                     #Source for 49th row
$sor492 = $sheetToCopy.Cells.Item(49,2)
$sor493 = $sheetToCopy.Cells.Item(49,3)
$sor494 = $sheetToCopy.Cells.Item(49,4)
$sor495 = $sheetToCopy.Cells.Item(49,5)
$sor496 = $sheetToCopy.Cells.Item(49,6)
$sor497 = $sheetToCopy.Cells.Item(49,7)           
$sor49last = $sheetToCopy.Cells.Item(49,8)

$sor501 = $sheetToCopy.Cells.Item(50,1)                     #Source for 50th row
$sor502 = $sheetToCopy.Cells.Item(50,2)
$sor503 = $sheetToCopy.Cells.Item(50,3)
$sor504 = $sheetToCopy.Cells.Item(50,4)
$sor505 = $sheetToCopy.Cells.Item(50,5)
$sor506 = $sheetToCopy.Cells.Item(50,6)
$sor507 = $sheetToCopy.Cells.Item(50,7)           
$sor50last = $sheetToCopy.Cells.Item(50,8)

$sor511 = $sheetToCopy.Cells.Item(51,1)                     #Source for 51th row
$sor512 = $sheetToCopy.Cells.Item(51,2)
$sor513 = $sheetToCopy.Cells.Item(51,3)
$sor514 = $sheetToCopy.Cells.Item(51,4)
$sor515 = $sheetToCopy.Cells.Item(51,5)
$sor516 = $sheetToCopy.Cells.Item(51,6)
$sor517 = $sheetToCopy.Cells.Item(51,7)           
$sor51last = $sheetToCopy.Cells.Item(51,8)

$sor521 = $sheetToCopy.Cells.Item(52,1)                     #Source for 52th row
$sor522 = $sheetToCopy.Cells.Item(52,2)
$sor523 = $sheetToCopy.Cells.Item(52,3)
$sor524 = $sheetToCopy.Cells.Item(52,4)
$sor525 = $sheetToCopy.Cells.Item(52,5)
$sor526 = $sheetToCopy.Cells.Item(52,6)
$sor527 = $sheetToCopy.Cells.Item(52,7)           
$sor52last = $sheetToCopy.Cells.Item(52,8)

$sor531 = $sheetToCopy.Cells.Item(53,1)                     #Source for 53th row
$sor532 = $sheetToCopy.Cells.Item(53,2)
$sor533 = $sheetToCopy.Cells.Item(53,3)
$sor534 = $sheetToCopy.Cells.Item(53,4)
$sor535 = $sheetToCopy.Cells.Item(53,5)
$sor536 = $sheetToCopy.Cells.Item(53,6)
$sor537 = $sheetToCopy.Cells.Item(53,7)           
$sor53last = $sheetToCopy.Cells.Item(53,8)

$sor541 = $sheetToCopy.Cells.Item(54,1)                     #Source for 54th row
$sor542 = $sheetToCopy.Cells.Item(54,2)
$sor543 = $sheetToCopy.Cells.Item(54,3)
$sor544 = $sheetToCopy.Cells.Item(54,4)
$sor545 = $sheetToCopy.Cells.Item(54,5)
$sor546 = $sheetToCopy.Cells.Item(54,6)
$sor547 = $sheetToCopy.Cells.Item(54,7)           
$sor54last = $sheetToCopy.Cells.Item(54,8)

$sor551 = $sheetToCopy.Cells.Item(55,1)                     #Source for 55th row
$sor552 = $sheetToCopy.Cells.Item(55,2)
$sor553 = $sheetToCopy.Cells.Item(55,3)
$sor554 = $sheetToCopy.Cells.Item(55,4)
$sor555 = $sheetToCopy.Cells.Item(55,5)
$sor556 = $sheetToCopy.Cells.Item(55,6)
$sor557 = $sheetToCopy.Cells.Item(55,7)           
$sor55last = $sheetToCopy.Cells.Item(55,8)

$sor561 = $sheetToCopy.Cells.Item(56,1)                     #Source for 56th row
$sor562 = $sheetToCopy.Cells.Item(56,2)
$sor563 = $sheetToCopy.Cells.Item(56,3)
$sor564 = $sheetToCopy.Cells.Item(56,4)
$sor565 = $sheetToCopy.Cells.Item(56,5)
$sor566 = $sheetToCopy.Cells.Item(56,6)
$sor567 = $sheetToCopy.Cells.Item(56,7)           
$sor56last = $sheetToCopy.Cells.Item(56,8)

$sor571 = $sheetToCopy.Cells.Item(57,1)                     #Source for 57th row
$sor572 = $sheetToCopy.Cells.Item(57,2)
$sor573 = $sheetToCopy.Cells.Item(57,3)
$sor574 = $sheetToCopy.Cells.Item(57,4)
$sor575 = $sheetToCopy.Cells.Item(57,5)
$sor576 = $sheetToCopy.Cells.Item(57,6)
$sor577 = $sheetToCopy.Cells.Item(57,7)           
$sor57last = $sheetToCopy.Cells.Item(57,8)

$sor581 = $sheetToCopy.Cells.Item(58,1)                     #Source for 58th row
$sor582 = $sheetToCopy.Cells.Item(58,2)
$sor583 = $sheetToCopy.Cells.Item(58,3)
$sor584 = $sheetToCopy.Cells.Item(58,4)
$sor585 = $sheetToCopy.Cells.Item(58,5)
$sor586 = $sheetToCopy.Cells.Item(58,6)
$sor587 = $sheetToCopy.Cells.Item(58,7)           
$sor58last = $sheetToCopy.Cells.Item(58,8)

$sor591 = $sheetToCopy.Cells.Item(59,1)                     #Source for 59th row
$sor592 = $sheetToCopy.Cells.Item(59,2)
$sor593 = $sheetToCopy.Cells.Item(59,3)
$sor594 = $sheetToCopy.Cells.Item(59,4)
$sor595 = $sheetToCopy.Cells.Item(59,5)
$sor596 = $sheetToCopy.Cells.Item(59,6)
$sor597 = $sheetToCopy.Cells.Item(59,7)           
$sor59last = $sheetToCopy.Cells.Item(59,8)

$sor601 = $sheetToCopy.Cells.Item(60,1)                     #Source for 60th row
$sor602 = $sheetToCopy.Cells.Item(60,2)
$sor603 = $sheetToCopy.Cells.Item(60,3)
$sor604 = $sheetToCopy.Cells.Item(60,4)
$sor605 = $sheetToCopy.Cells.Item(60,5)
$sor606 = $sheetToCopy.Cells.Item(60,6)
$sor607 = $sheetToCopy.Cells.Item(60,7)           
$sor60last = $sheetToCopy.Cells.Item(60,8)

$sor611 = $sheetToCopy.Cells.Item(61,1)                     #Source for 61th row
$sor612 = $sheetToCopy.Cells.Item(61,2)
$sor613 = $sheetToCopy.Cells.Item(61,3)
$sor614 = $sheetToCopy.Cells.Item(61,4)
$sor615 = $sheetToCopy.Cells.Item(61,5)
$sor616 = $sheetToCopy.Cells.Item(61,6)
$sor617 = $sheetToCopy.Cells.Item(61,7)           
$sor61last = $sheetToCopy.Cells.Item(61,8)

$sor621 = $sheetToCopy.Cells.Item(62,1)                     #Source for 62th row
$sor622 = $sheetToCopy.Cells.Item(62,2)
$sor623 = $sheetToCopy.Cells.Item(62,3)
$sor624 = $sheetToCopy.Cells.Item(62,4)
$sor625 = $sheetToCopy.Cells.Item(62,5)
$sor626 = $sheetToCopy.Cells.Item(62,6)
$sor627 = $sheetToCopy.Cells.Item(62,7)           
$sor62last = $sheetToCopy.Cells.Item(62,8)

$sor631 = $sheetToCopy.Cells.Item(63,1)                     #Source for 63th row
$sor632 = $sheetToCopy.Cells.Item(63,2)
$sor633 = $sheetToCopy.Cells.Item(63,3)
$sor634 = $sheetToCopy.Cells.Item(63,4)
$sor635 = $sheetToCopy.Cells.Item(63,5)
$sor636 = $sheetToCopy.Cells.Item(63,6)
$sor637 = $sheetToCopy.Cells.Item(63,7)           
$sor63last = $sheetToCopy.Cells.Item(63,8)

$sor641 = $sheetToCopy.Cells.Item(64,1)                     #Source for 64th row
$sor642 = $sheetToCopy.Cells.Item(64,2)
$sor643 = $sheetToCopy.Cells.Item(64,3)
$sor644 = $sheetToCopy.Cells.Item(64,4)
$sor645 = $sheetToCopy.Cells.Item(64,5)
$sor646 = $sheetToCopy.Cells.Item(64,6)
$sor647 = $sheetToCopy.Cells.Item(64,7)           
$sor64last = $sheetToCopy.Cells.Item(64,8)

$sor651 = $sheetToCopy.Cells.Item(65,1)                     #Source for 65th row
$sor652 = $sheetToCopy.Cells.Item(65,2)
$sor653 = $sheetToCopy.Cells.Item(65,3)
$sor654 = $sheetToCopy.Cells.Item(65,4)
$sor655 = $sheetToCopy.Cells.Item(65,5)
$sor656 = $sheetToCopy.Cells.Item(65,6)
$sor657 = $sheetToCopy.Cells.Item(65,7)           
$sor65last = $sheetToCopy.Cells.Item(65,8)

$sor661 = $sheetToCopy.Cells.Item(66,1)                     #Source for 66th row
$sor662 = $sheetToCopy.Cells.Item(66,2)
$sor663 = $sheetToCopy.Cells.Item(66,3)
$sor664 = $sheetToCopy.Cells.Item(66,4)
$sor665 = $sheetToCopy.Cells.Item(66,5)
$sor666 = $sheetToCopy.Cells.Item(66,6)
$sor667 = $sheetToCopy.Cells.Item(66,7)           
$sor66last = $sheetToCopy.Cells.Item(66,8)

$sor671 = $sheetToCopy.Cells.Item(67,1)                     #Source for 67th row
$sor672 = $sheetToCopy.Cells.Item(67,2)
$sor673 = $sheetToCopy.Cells.Item(67,3)
$sor674 = $sheetToCopy.Cells.Item(67,4)
$sor675 = $sheetToCopy.Cells.Item(67,5)
$sor676 = $sheetToCopy.Cells.Item(67,6)
$sor677 = $sheetToCopy.Cells.Item(67,7)           
$sor67last = $sheetToCopy.Cells.Item(67,8)

$sor681 = $sheetToCopy.Cells.Item(68,1)                     #Source for 68th row
$sor682 = $sheetToCopy.Cells.Item(68,2)
$sor683 = $sheetToCopy.Cells.Item(68,3)
$sor684 = $sheetToCopy.Cells.Item(68,4)
$sor685 = $sheetToCopy.Cells.Item(68,5)
$sor686 = $sheetToCopy.Cells.Item(68,6)
$sor687 = $sheetToCopy.Cells.Item(68,7)           
$sor68last = $sheetToCopy.Cells.Item(68,8)

$sor691 = $sheetToCopy.Cells.Item(69,1)                     #Source for 69th row
$sor692 = $sheetToCopy.Cells.Item(69,2)
$sor693 = $sheetToCopy.Cells.Item(69,3)
$sor694 = $sheetToCopy.Cells.Item(69,4)
$sor695 = $sheetToCopy.Cells.Item(69,5)
$sor696 = $sheetToCopy.Cells.Item(69,6)
$sor697 = $sheetToCopy.Cells.Item(69,7)           
$sor69last = $sheetToCopy.Cells.Item(69,8)

$sor701 = $sheetToCopy.Cells.Item(70,1)                     #Source for 70th row
$sor702 = $sheetToCopy.Cells.Item(70,2)
$sor703 = $sheetToCopy.Cells.Item(70,3)
$sor704 = $sheetToCopy.Cells.Item(70,4)
$sor705 = $sheetToCopy.Cells.Item(70,5)
$sor706 = $sheetToCopy.Cells.Item(70,6)
$sor707 = $sheetToCopy.Cells.Item(70,7)           
$sor70last = $sheetToCopy.Cells.Item(70,8)

$sor711 = $sheetToCopy.Cells.Item(71,1)                     #Source for 71th row
$sor712 = $sheetToCopy.Cells.Item(71,2)
$sor713 = $sheetToCopy.Cells.Item(71,3)
$sor714 = $sheetToCopy.Cells.Item(71,4)
$sor715 = $sheetToCopy.Cells.Item(71,5)
$sor716 = $sheetToCopy.Cells.Item(71,6)
$sor717 = $sheetToCopy.Cells.Item(71,7)           
$sor71last = $sheetToCopy.Cells.Item(71,8)

$sor721 = $sheetToCopy.Cells.Item(72,1)                     #Source for 72th row
$sor722 = $sheetToCopy.Cells.Item(72,2)
$sor723 = $sheetToCopy.Cells.Item(72,3)
$sor724 = $sheetToCopy.Cells.Item(72,4)
$sor725 = $sheetToCopy.Cells.Item(72,5)
$sor726 = $sheetToCopy.Cells.Item(72,6)
$sor727 = $sheetToCopy.Cells.Item(72,7)           
$sor72last = $sheetToCopy.Cells.Item(72,8)

$sor731 = $sheetToCopy.Cells.Item(73,1)                     #Source for 73th row
$sor732 = $sheetToCopy.Cells.Item(73,2)
$sor733 = $sheetToCopy.Cells.Item(73,3)
$sor734 = $sheetToCopy.Cells.Item(73,4)
$sor735 = $sheetToCopy.Cells.Item(73,5)
$sor736 = $sheetToCopy.Cells.Item(73,6)
$sor737 = $sheetToCopy.Cells.Item(73,7)           
$sor73last = $sheetToCopy.Cells.Item(73,8)

$sor741 = $sheetToCopy.Cells.Item(74,1)                     #Source for 74th row
$sor742 = $sheetToCopy.Cells.Item(74,2)
$sor743 = $sheetToCopy.Cells.Item(74,3)
$sor744 = $sheetToCopy.Cells.Item(74,4)
$sor745 = $sheetToCopy.Cells.Item(74,5)
$sor746 = $sheetToCopy.Cells.Item(74,6)
$sor747 = $sheetToCopy.Cells.Item(74,7)           
$sor74last = $sheetToCopy.Cells.Item(74,8)

$sor751 = $sheetToCopy.Cells.Item(75,1)                     #Source for 75th row
$sor752 = $sheetToCopy.Cells.Item(75,2)
$sor753 = $sheetToCopy.Cells.Item(75,3)
$sor754 = $sheetToCopy.Cells.Item(75,4)
$sor755 = $sheetToCopy.Cells.Item(75,5)
$sor756 = $sheetToCopy.Cells.Item(75,6)
$sor757 = $sheetToCopy.Cells.Item(75,7)           
$sor75last = $sheetToCopy.Cells.Item(75,8)

$sor761 = $sheetToCopy.Cells.Item(76,1)                     #Source for 76th row
$sor762 = $sheetToCopy.Cells.Item(76,2)
$sor763 = $sheetToCopy.Cells.Item(76,3)
$sor764 = $sheetToCopy.Cells.Item(76,4)
$sor765 = $sheetToCopy.Cells.Item(76,5)
$sor766 = $sheetToCopy.Cells.Item(76,6)
$sor767 = $sheetToCopy.Cells.Item(76,7)           
$sor76last = $sheetToCopy.Cells.Item(76,8)

$sor771 = $sheetToCopy.Cells.Item(77,1)                     #Source for 77th row
$sor772 = $sheetToCopy.Cells.Item(77,2)
$sor773 = $sheetToCopy.Cells.Item(77,3)
$sor774 = $sheetToCopy.Cells.Item(77,4)
$sor775 = $sheetToCopy.Cells.Item(77,5)
$sor776 = $sheetToCopy.Cells.Item(77,6)
$sor777 = $sheetToCopy.Cells.Item(77,7)           
$sor77last = $sheetToCopy.Cells.Item(77,8)

$sor781 = $sheetToCopy.Cells.Item(78,1)                     #Source for 78th row
$sor782 = $sheetToCopy.Cells.Item(78,2)
$sor783 = $sheetToCopy.Cells.Item(78,3)
$sor784 = $sheetToCopy.Cells.Item(78,4)
$sor785 = $sheetToCopy.Cells.Item(78,5)
$sor786 = $sheetToCopy.Cells.Item(78,6)
$sor787 = $sheetToCopy.Cells.Item(78,7)           
$sor78last = $sheetToCopy.Cells.Item(78,8)

$sor791 = $sheetToCopy.Cells.Item(79,1)                     #Source for 79th row
$sor792 = $sheetToCopy.Cells.Item(79,2)
$sor793 = $sheetToCopy.Cells.Item(79,3)
$sor794 = $sheetToCopy.Cells.Item(79,4)
$sor795 = $sheetToCopy.Cells.Item(79,5)
$sor796 = $sheetToCopy.Cells.Item(79,6)
$sor797 = $sheetToCopy.Cells.Item(79,7)           
$sor79last = $sheetToCopy.Cells.Item(79,8)

$sor801 = $sheetToCopy.Cells.Item(80,1)                     #Source for 80th row
$sor802 = $sheetToCopy.Cells.Item(80,2)
$sor803 = $sheetToCopy.Cells.Item(80,3)
$sor804 = $sheetToCopy.Cells.Item(80,4)
$sor805 = $sheetToCopy.Cells.Item(80,5)
$sor806 = $sheetToCopy.Cells.Item(80,6)
$sor807 = $sheetToCopy.Cells.Item(80,7)           
$sor80last = $sheetToCopy.Cells.Item(80,8)

$sor811 = $sheetToCopy.Cells.Item(81,1)                     #Source for 81th row
$sor812 = $sheetToCopy.Cells.Item(81,2)
$sor813 = $sheetToCopy.Cells.Item(81,3)
$sor814 = $sheetToCopy.Cells.Item(81,4)
$sor815 = $sheetToCopy.Cells.Item(81,5)
$sor816 = $sheetToCopy.Cells.Item(81,6)
$sor817 = $sheetToCopy.Cells.Item(81,7)           
$sor81last = $sheetToCopy.Cells.Item(81,8)

$sor821 = $sheetToCopy.Cells.Item(82,1)                     #Source for 82th row
$sor822 = $sheetToCopy.Cells.Item(82,2)
$sor823 = $sheetToCopy.Cells.Item(82,3)
$sor824 = $sheetToCopy.Cells.Item(82,4)
$sor825 = $sheetToCopy.Cells.Item(82,5)
$sor826 = $sheetToCopy.Cells.Item(82,6)
$sor827 = $sheetToCopy.Cells.Item(82,7)           
$sor82last = $sheetToCopy.Cells.Item(82,8)

$sor831 = $sheetToCopy.Cells.Item(83,1)                     #Source for 83th row
$sor832 = $sheetToCopy.Cells.Item(83,2)
$sor833 = $sheetToCopy.Cells.Item(83,3)
$sor834 = $sheetToCopy.Cells.Item(83,4)
$sor835 = $sheetToCopy.Cells.Item(83,5)
$sor836 = $sheetToCopy.Cells.Item(83,6)
$sor837 = $sheetToCopy.Cells.Item(83,7)           
$sor83last = $sheetToCopy.Cells.Item(83,8)

$sor841 = $sheetToCopy.Cells.Item(84,1)                     #Source for 84th row
$sor842 = $sheetToCopy.Cells.Item(84,2)
$sor843 = $sheetToCopy.Cells.Item(84,3)
$sor844 = $sheetToCopy.Cells.Item(84,4)
$sor845 = $sheetToCopy.Cells.Item(84,5)
$sor846 = $sheetToCopy.Cells.Item(84,6)
$sor847 = $sheetToCopy.Cells.Item(84,7)           
$sor84last = $sheetToCopy.Cells.Item(84,8)

$sor851 = $sheetToCopy.Cells.Item(85,1)                     #Source for 85th row
$sor852 = $sheetToCopy.Cells.Item(85,2)
$sor853 = $sheetToCopy.Cells.Item(85,3)
$sor854 = $sheetToCopy.Cells.Item(85,4)
$sor855 = $sheetToCopy.Cells.Item(85,5)
$sor856 = $sheetToCopy.Cells.Item(85,6)
$sor857 = $sheetToCopy.Cells.Item(85,7)           
$sor85last = $sheetToCopy.Cells.Item(85,8)

$sor861 = $sheetToCopy.Cells.Item(86,1)                     #Source for 86th row
$sor862 = $sheetToCopy.Cells.Item(86,2)
$sor863 = $sheetToCopy.Cells.Item(86,3)
$sor864 = $sheetToCopy.Cells.Item(86,4)
$sor865 = $sheetToCopy.Cells.Item(86,5)
$sor866 = $sheetToCopy.Cells.Item(86,6)
$sor867 = $sheetToCopy.Cells.Item(86,7)           
$sor86last = $sheetToCopy.Cells.Item(86,8)

$sor871 = $sheetToCopy.Cells.Item(87,1)                     #Source for 87th row
$sor872 = $sheetToCopy.Cells.Item(87,2)
$sor873 = $sheetToCopy.Cells.Item(87,3)
$sor874 = $sheetToCopy.Cells.Item(87,4)
$sor875 = $sheetToCopy.Cells.Item(87,5)
$sor876 = $sheetToCopy.Cells.Item(87,6)
$sor877 = $sheetToCopy.Cells.Item(87,7)           
$sor87last = $sheetToCopy.Cells.Item(87,8)

$sor881 = $sheetToCopy.Cells.Item(88,1)                     #Source for 88th row
$sor882 = $sheetToCopy.Cells.Item(88,2)
$sor883 = $sheetToCopy.Cells.Item(88,3)
$sor884 = $sheetToCopy.Cells.Item(88,4)
$sor885 = $sheetToCopy.Cells.Item(88,5)
$sor886 = $sheetToCopy.Cells.Item(88,6)
$sor887 = $sheetToCopy.Cells.Item(88,7)           
$sor88last = $sheetToCopy.Cells.Item(88,8)

$sor891 = $sheetToCopy.Cells.Item(89,1)                     #Source for 89th row
$sor892 = $sheetToCopy.Cells.Item(89,2)
$sor893 = $sheetToCopy.Cells.Item(89,3)
$sor894 = $sheetToCopy.Cells.Item(89,4)
$sor895 = $sheetToCopy.Cells.Item(89,5)
$sor896 = $sheetToCopy.Cells.Item(89,6)
$sor897 = $sheetToCopy.Cells.Item(89,7)           
$sor89last = $sheetToCopy.Cells.Item(89,8)

$sor901 = $sheetToCopy.Cells.Item(90,1)                     #Source for 90th row
$sor902 = $sheetToCopy.Cells.Item(90,2)
$sor903 = $sheetToCopy.Cells.Item(90,3)
$sor904 = $sheetToCopy.Cells.Item(90,4)
$sor905 = $sheetToCopy.Cells.Item(90,5)
$sor906 = $sheetToCopy.Cells.Item(90,6)
$sor907 = $sheetToCopy.Cells.Item(90,7)           
$sor90last = $sheetToCopy.Cells.Item(90,8)

$sor911 = $sheetToCopy.Cells.Item(91,1)                     #Source for 91th row
$sor912 = $sheetToCopy.Cells.Item(91,2)
$sor913 = $sheetToCopy.Cells.Item(91,3)
$sor914 = $sheetToCopy.Cells.Item(91,4)
$sor915 = $sheetToCopy.Cells.Item(91,5)
$sor916 = $sheetToCopy.Cells.Item(91,6)
$sor917 = $sheetToCopy.Cells.Item(91,7)           
$sor91last = $sheetToCopy.Cells.Item(91,8)

$sor921 = $sheetToCopy.Cells.Item(92,1)                     #Source for 92th row
$sor922 = $sheetToCopy.Cells.Item(92,2)
$sor923 = $sheetToCopy.Cells.Item(92,3)
$sor924 = $sheetToCopy.Cells.Item(92,4)
$sor925 = $sheetToCopy.Cells.Item(92,5)
$sor926 = $sheetToCopy.Cells.Item(92,6)
$sor927 = $sheetToCopy.Cells.Item(92,7)           
$sor92last = $sheetToCopy.Cells.Item(92,8)

$sor931 = $sheetToCopy.Cells.Item(93,1)                     #Source for 93th row
$sor932 = $sheetToCopy.Cells.Item(93,2)
$sor933 = $sheetToCopy.Cells.Item(93,3)
$sor934 = $sheetToCopy.Cells.Item(93,4)
$sor935 = $sheetToCopy.Cells.Item(93,5)
$sor936 = $sheetToCopy.Cells.Item(93,6)
$sor937 = $sheetToCopy.Cells.Item(93,7)           
$sor93last = $sheetToCopy.Cells.Item(93,8)

$sor941 = $sheetToCopy.Cells.Item(94,1)                     #Source for 94th row
$sor942 = $sheetToCopy.Cells.Item(94,2)
$sor943 = $sheetToCopy.Cells.Item(94,3)
$sor944 = $sheetToCopy.Cells.Item(94,4)
$sor945 = $sheetToCopy.Cells.Item(94,5)
$sor946 = $sheetToCopy.Cells.Item(94,6)
$sor947 = $sheetToCopy.Cells.Item(94,7)           
$sor94last = $sheetToCopy.Cells.Item(94,8)

$sor951 = $sheetToCopy.Cells.Item(95,1)                     #Source for 95th row
$sor952 = $sheetToCopy.Cells.Item(95,2)
$sor953 = $sheetToCopy.Cells.Item(95,3)
$sor954 = $sheetToCopy.Cells.Item(95,4)
$sor955 = $sheetToCopy.Cells.Item(95,5)
$sor956 = $sheetToCopy.Cells.Item(95,6)
$sor957 = $sheetToCopy.Cells.Item(95,7)           
$sor95last = $sheetToCopy.Cells.Item(95,8)

$sor961 = $sheetToCopy.Cells.Item(96,1)                     #Source for 96th row
$sor962 = $sheetToCopy.Cells.Item(96,2)
$sor963 = $sheetToCopy.Cells.Item(96,3)
$sor964 = $sheetToCopy.Cells.Item(96,4)
$sor965 = $sheetToCopy.Cells.Item(96,5)
$sor966 = $sheetToCopy.Cells.Item(96,6)
$sor967 = $sheetToCopy.Cells.Item(96,7)           
$sor96last = $sheetToCopy.Cells.Item(96,8)

$sor971 = $sheetToCopy.Cells.Item(97,1)                     #Source for 97th row
$sor972 = $sheetToCopy.Cells.Item(97,2)
$sor973 = $sheetToCopy.Cells.Item(97,3)
$sor974 = $sheetToCopy.Cells.Item(97,4)
$sor975 = $sheetToCopy.Cells.Item(97,5)
$sor976 = $sheetToCopy.Cells.Item(97,6)
$sor977 = $sheetToCopy.Cells.Item(97,7)           
$sor97last = $sheetToCopy.Cells.Item(97,8)

$sor981 = $sheetToCopy.Cells.Item(98,1)                     #Source for 98th row
$sor982 = $sheetToCopy.Cells.Item(98,2)
$sor983 = $sheetToCopy.Cells.Item(98,3)
$sor984 = $sheetToCopy.Cells.Item(98,4)
$sor985 = $sheetToCopy.Cells.Item(98,5)
$sor986 = $sheetToCopy.Cells.Item(98,6)
$sor987 = $sheetToCopy.Cells.Item(98,7)           
$sor98last = $sheetToCopy.Cells.Item(98,8)

$sor991 = $sheetToCopy.Cells.Item(99,1)                     #Source for 99th row
$sor992 = $sheetToCopy.Cells.Item(99,2)
$sor993 = $sheetToCopy.Cells.Item(99,3)
$sor994 = $sheetToCopy.Cells.Item(99,4)
$sor995 = $sheetToCopy.Cells.Item(99,5)
$sor996 = $sheetToCopy.Cells.Item(99,6)
$sor997 = $sheetToCopy.Cells.Item(99,7)           
$sor99last = $sheetToCopy.Cells.Item(99,8)

$sor1001 = $sheetToCopy.Cells.Item(100,1)                     #Source for 100th row
$sor1002 = $sheetToCopy.Cells.Item(100,2)
$sor1003 = $sheetToCopy.Cells.Item(100,3)
$sor1004 = $sheetToCopy.Cells.Item(100,4)
$sor1005 = $sheetToCopy.Cells.Item(100,5)
$sor1006 = $sheetToCopy.Cells.Item(100,6)
$sor1007 = $sheetToCopy.Cells.Item(100,7)           
$sor100last = $sheetToCopy.Cells.Item(100,8)
#--------------------------------------------------------------------------------------------

Add-Type -AssemblyName Microsoft.Office.Interop.Excel
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
$xl.Visible = $true


$ws = $wb.Worksheets.Item(1)


$objRange = $ws.UsedRange
$a = $objRange.SpecialCells(11).row
$b = $objRange.SpecialCells(11).column
write-host "Lastrow:", $a, " Last Column:" $b

#------------------------------------------------------------------------------------------

$tar1 = $ws.Cells.Item($a,1)                              #Target 
$tar2 = $ws.Cells.Item($a,2)
$tar3 = $ws.Cells.Item($a,3)
$tar4 = $ws.Cells.Item($a,4)
$tar5 = $ws.Cells.Item($a,5)
$tar6 = $ws.Cells.Item($a,6)
$tar7 = $ws.Cells.Item($a,7)                         
$tarlast = $ws.Cells.Item($a,$b)

$tar21 = $ws.Cells.Item(($a+1),1)                         #Target for second row
$tar22 = $ws.Cells.Item(($a+1),2)
$tar23 = $ws.Cells.Item(($a+1),3)
$tar24 = $ws.Cells.Item(($a+1),4)
$tar25 = $ws.Cells.Item(($a+1),5)
$tar26 = $ws.Cells.Item(($a+1),6)
$tar27 = $ws.Cells.Item(($a+1),7)                         
$tar2last = $ws.Cells.Item(($a+1),$b) 

$tar31 = $ws.Cells.Item(($a+2),1)                         #Target for 3th row
$tar32 = $ws.Cells.Item(($a+2),2)
$tar33 = $ws.Cells.Item(($a+2),3)
$tar34 = $ws.Cells.Item(($a+2),4)
$tar35 = $ws.Cells.Item(($a+2),5)
$tar36 = $ws.Cells.Item(($a+2),6)
$tar37 = $ws.Cells.Item(($a+2),7)                         
$tar3last = $ws.Cells.Item(($a+2),$b)

$tar41 = $ws.Cells.Item(($a+3),1)                         #Target for 4th row
$tar42 = $ws.Cells.Item(($a+3),2)
$tar43 = $ws.Cells.Item(($a+3),3)
$tar44 = $ws.Cells.Item(($a+3),4)
$tar45 = $ws.Cells.Item(($a+3),5)
$tar46 = $ws.Cells.Item(($a+3),6)
$tar47 = $ws.Cells.Item(($a+3),7)                         
$tar4last = $ws.Cells.Item(($a+3),$b)

$tar51 = $ws.Cells.Item(($a+4),1)                         #Target for 5th row
$tar52 = $ws.Cells.Item(($a+4),2)
$tar53 = $ws.Cells.Item(($a+4),3)
$tar54 = $ws.Cells.Item(($a+4),4)
$tar55 = $ws.Cells.Item(($a+4),5)
$tar56 = $ws.Cells.Item(($a+4),6)
$tar57 = $ws.Cells.Item(($a+4),7)                         
$tar5last = $ws.Cells.Item(($a+4),$b)

$tar61 = $ws.Cells.Item(($a+5),1)                         #Target for 6th row
$tar62 = $ws.Cells.Item(($a+5),2)
$tar63 = $ws.Cells.Item(($a+5),3)
$tar64 = $ws.Cells.Item(($a+5),4)
$tar65 = $ws.Cells.Item(($a+5),5)
$tar66 = $ws.Cells.Item(($a+5),6)
$tar67 = $ws.Cells.Item(($a+5),7)                         
$tar6last = $ws.Cells.Item(($a+5),$b)

$tar71 = $ws.Cells.Item(($a+6),1)                         #Target for 7th row
$tar72 = $ws.Cells.Item(($a+6),2)
$tar73 = $ws.Cells.Item(($a+6),3)
$tar74 = $ws.Cells.Item(($a+6),4)
$tar75 = $ws.Cells.Item(($a+6),5)
$tar76 = $ws.Cells.Item(($a+6),6)
$tar77 = $ws.Cells.Item(($a+6),7)                         
$tar7last = $ws.Cells.Item(($a+6),$b)

$tar81 = $ws.Cells.Item(($a+7),1)                         #Target for 8th row
$tar82 = $ws.Cells.Item(($a+7),2)
$tar83 = $ws.Cells.Item(($a+7),3)
$tar84 = $ws.Cells.Item(($a+7),4)
$tar85 = $ws.Cells.Item(($a+7),5)
$tar86 = $ws.Cells.Item(($a+7),6)
$tar87 = $ws.Cells.Item(($a+7),7)                         
$tar8last = $ws.Cells.Item(($a+7),$b)

$tar91 = $ws.Cells.Item(($a+8),1)                         #Target for 9th row
$tar92 = $ws.Cells.Item(($a+8),2)
$tar93 = $ws.Cells.Item(($a+8),3)
$tar94 = $ws.Cells.Item(($a+8),4)
$tar95 = $ws.Cells.Item(($a+8),5)
$tar96 = $ws.Cells.Item(($a+8),6)
$tar97 = $ws.Cells.Item(($a+8),7)                         
$tar9last = $ws.Cells.Item(($a+8),$b)

$tar101 = $ws.Cells.Item(($a+9),1)                         #Target for 10th row
$tar102 = $ws.Cells.Item(($a+9),2)
$tar103 = $ws.Cells.Item(($a+9),3)
$tar104 = $ws.Cells.Item(($a+9),4)
$tar105 = $ws.Cells.Item(($a+9),5)
$tar106 = $ws.Cells.Item(($a+9),6)
$tar107 = $ws.Cells.Item(($a+9),7)                         
$tar10last = $ws.Cells.Item(($a+9),$b)

$tar111 = $ws.Cells.Item(($a+10),1)                         #Target for 11th row
$tar112 = $ws.Cells.Item(($a+10),2)
$tar113 = $ws.Cells.Item(($a+10),3)
$tar114 = $ws.Cells.Item(($a+10),4)
$tar115 = $ws.Cells.Item(($a+10),5)
$tar116 = $ws.Cells.Item(($a+10),6)
$tar117 = $ws.Cells.Item(($a+10),7)                         
$tar11last = $ws.Cells.Item(($a+10),$b)

$tar121 = $ws.Cells.Item(($a+11),1)                         #Target for 12th row
$tar122 = $ws.Cells.Item(($a+11),2)
$tar123 = $ws.Cells.Item(($a+11),3)
$tar124 = $ws.Cells.Item(($a+11),4)
$tar125 = $ws.Cells.Item(($a+11),5)
$tar126 = $ws.Cells.Item(($a+11),6)
$tar127 = $ws.Cells.Item(($a+11),7)                         
$tar12last = $ws.Cells.Item(($a+11),$b)

$tar131 = $ws.Cells.Item(($a+12),1)                         #Target for 13th row
$tar132 = $ws.Cells.Item(($a+12),2)
$tar133 = $ws.Cells.Item(($a+12),3)
$tar134 = $ws.Cells.Item(($a+12),4)
$tar135 = $ws.Cells.Item(($a+12),5)
$tar136 = $ws.Cells.Item(($a+12),6)
$tar137 = $ws.Cells.Item(($a+12),7)                         
$tar13last = $ws.Cells.Item(($a+12),$b)

$tar141 = $ws.Cells.Item(($a+13),1)                         #Target for 14th row
$tar142 = $ws.Cells.Item(($a+13),2)
$tar143 = $ws.Cells.Item(($a+13),3)
$tar144 = $ws.Cells.Item(($a+13),4)
$tar145 = $ws.Cells.Item(($a+13),5)
$tar146 = $ws.Cells.Item(($a+13),6)
$tar147 = $ws.Cells.Item(($a+13),7)                         
$tar14last = $ws.Cells.Item(($a+13),$b)

$tar151 = $ws.Cells.Item(($a+14),1)                         #Target for 15th row
$tar152 = $ws.Cells.Item(($a+14),2)
$tar153 = $ws.Cells.Item(($a+14),3)
$tar154 = $ws.Cells.Item(($a+14),4)
$tar155 = $ws.Cells.Item(($a+14),5)
$tar156 = $ws.Cells.Item(($a+14),6)
$tar157 = $ws.Cells.Item(($a+14),7)                         
$tar15last = $ws.Cells.Item(($a+14),$b)

$tar161 = $ws.Cells.Item(($a+15),1)                         #Target for 16th row
$tar162 = $ws.Cells.Item(($a+15),2)
$tar163 = $ws.Cells.Item(($a+15),3)
$tar164 = $ws.Cells.Item(($a+15),4)
$tar165 = $ws.Cells.Item(($a+15),5)
$tar166 = $ws.Cells.Item(($a+15),6)
$tar167 = $ws.Cells.Item(($a+15),7)                         
$tar16last = $ws.Cells.Item(($a+15),$b)

$tar171 = $ws.Cells.Item(($a+16),1)                         #Target for 17th row
$tar172 = $ws.Cells.Item(($a+16),2)
$tar173 = $ws.Cells.Item(($a+16),3)
$tar174 = $ws.Cells.Item(($a+16),4)
$tar175 = $ws.Cells.Item(($a+16),5)
$tar176 = $ws.Cells.Item(($a+16),6)
$tar177 = $ws.Cells.Item(($a+16),7)                         
$tar17last = $ws.Cells.Item(($a+16),$b)

$tar181 = $ws.Cells.Item(($a+17),1)                         #Target for 18th row
$tar182 = $ws.Cells.Item(($a+17),2)
$tar183 = $ws.Cells.Item(($a+17),3)
$tar184 = $ws.Cells.Item(($a+17),4)
$tar185 = $ws.Cells.Item(($a+17),5)
$tar186 = $ws.Cells.Item(($a+17),6)
$tar187 = $ws.Cells.Item(($a+17),7)                         
$tar18last = $ws.Cells.Item(($a+17),$b)

$tar191 = $ws.Cells.Item(($a+18),1)                         #Target for 19th row
$tar192 = $ws.Cells.Item(($a+18),2)
$tar193 = $ws.Cells.Item(($a+18),3)
$tar194 = $ws.Cells.Item(($a+18),4)
$tar195 = $ws.Cells.Item(($a+18),5)
$tar196 = $ws.Cells.Item(($a+18),6)
$tar197 = $ws.Cells.Item(($a+18),7)                         
$tar19last = $ws.Cells.Item(($a+18),$b)

$tar201 = $ws.Cells.Item(($a+19),1)                         #Target for 20th row
$tar202 = $ws.Cells.Item(($a+19),2)
$tar203 = $ws.Cells.Item(($a+19),3)
$tar204 = $ws.Cells.Item(($a+19),4)
$tar205 = $ws.Cells.Item(($a+19),5)
$tar206 = $ws.Cells.Item(($a+19),6)
$tar207 = $ws.Cells.Item(($a+19),7)                         
$tar20last = $ws.Cells.Item(($a+19),$b)

$tar211 = $ws.Cells.Item(($a+20),1)                         #Target for 21th row
$tar212 = $ws.Cells.Item(($a+20),2)
$tar213 = $ws.Cells.Item(($a+20),3)
$tar214 = $ws.Cells.Item(($a+20),4)
$tar215 = $ws.Cells.Item(($a+20),5)
$tar216 = $ws.Cells.Item(($a+20),6)
$tar217 = $ws.Cells.Item(($a+20),7)                         
$tar21last = $ws.Cells.Item(($a+20),$b)

$tar221 = $ws.Cells.Item(($a+21),1)                         #Target for 22th row
$tar222 = $ws.Cells.Item(($a+21),2)
$tar223 = $ws.Cells.Item(($a+21),3)
$tar224 = $ws.Cells.Item(($a+21),4)
$tar225 = $ws.Cells.Item(($a+21),5)
$tar226 = $ws.Cells.Item(($a+21),6)
$tar227 = $ws.Cells.Item(($a+21),7)                         
$tar22last = $ws.Cells.Item(($a+21),$b)

$tar231 = $ws.Cells.Item(($a+22),1)                         #Target for 23th row
$tar232 = $ws.Cells.Item(($a+22),2)
$tar233 = $ws.Cells.Item(($a+22),3)
$tar234 = $ws.Cells.Item(($a+22),4)
$tar235 = $ws.Cells.Item(($a+22),5)
$tar236 = $ws.Cells.Item(($a+22),6)
$tar237 = $ws.Cells.Item(($a+22),7)                         
$tar23last = $ws.Cells.Item(($a+22),$b)

$tar241 = $ws.Cells.Item(($a+23),1)                         #Target for 24th row
$tar242 = $ws.Cells.Item(($a+23),2)
$tar243 = $ws.Cells.Item(($a+23),3)
$tar244 = $ws.Cells.Item(($a+23),4)
$tar245 = $ws.Cells.Item(($a+23),5)
$tar246 = $ws.Cells.Item(($a+23),6)
$tar247 = $ws.Cells.Item(($a+23),7)                         
$tar24last = $ws.Cells.Item(($a+23),$b)

$tar251 = $ws.Cells.Item(($a+24),1)                         #Target for 25th row
$tar252 = $ws.Cells.Item(($a+24),2)
$tar253 = $ws.Cells.Item(($a+24),3)
$tar254 = $ws.Cells.Item(($a+24),4)
$tar255 = $ws.Cells.Item(($a+24),5)
$tar256 = $ws.Cells.Item(($a+24),6)
$tar257 = $ws.Cells.Item(($a+24),7)                         
$tar25last = $ws.Cells.Item(($a+24),$b)

$tar261 = $ws.Cells.Item(($a+25),1)                         #Target for 26th row
$tar262 = $ws.Cells.Item(($a+25),2)
$tar263 = $ws.Cells.Item(($a+25),3)
$tar264 = $ws.Cells.Item(($a+25),4)
$tar265 = $ws.Cells.Item(($a+25),5)
$tar266 = $ws.Cells.Item(($a+25),6)
$tar267 = $ws.Cells.Item(($a+25),7)                         
$tar26last = $ws.Cells.Item(($a+25),$b)

$tar271 = $ws.Cells.Item(($a+26),1)                         #Target for 27th row
$tar272 = $ws.Cells.Item(($a+26),2)
$tar273 = $ws.Cells.Item(($a+26),3)
$tar274 = $ws.Cells.Item(($a+26),4)
$tar275 = $ws.Cells.Item(($a+26),5)
$tar276 = $ws.Cells.Item(($a+26),6)
$tar277 = $ws.Cells.Item(($a+26),7)                         
$tar27last = $ws.Cells.Item(($a+26),$b)

$tar281 = $ws.Cells.Item(($a+27),1)                         #Target for 28th row
$tar282 = $ws.Cells.Item(($a+27),2)
$tar283 = $ws.Cells.Item(($a+27),3)
$tar284 = $ws.Cells.Item(($a+27),4)
$tar285 = $ws.Cells.Item(($a+27),5)
$tar286 = $ws.Cells.Item(($a+27),6)
$tar287 = $ws.Cells.Item(($a+27),7)                         
$tar28last = $ws.Cells.Item(($a+27),$b)

$tar291 = $ws.Cells.Item(($a+28),1)                         #Target for 29th row
$tar292 = $ws.Cells.Item(($a+28),2)
$tar293 = $ws.Cells.Item(($a+28),3)
$tar294 = $ws.Cells.Item(($a+28),4)
$tar295 = $ws.Cells.Item(($a+28),5)
$tar296 = $ws.Cells.Item(($a+28),6)
$tar297 = $ws.Cells.Item(($a+28),7)                         
$tar29last = $ws.Cells.Item(($a+28),$b)

$tar301 = $ws.Cells.Item(($a+29),1)                         #Target for 30th row
$tar302 = $ws.Cells.Item(($a+29),2)
$tar303 = $ws.Cells.Item(($a+29),3)
$tar304 = $ws.Cells.Item(($a+29),4)
$tar305 = $ws.Cells.Item(($a+29),5)
$tar306 = $ws.Cells.Item(($a+29),6)
$tar307 = $ws.Cells.Item(($a+29),7)                         
$tar30last = $ws.Cells.Item(($a+29),$b)

$tar311 = $ws.Cells.Item(($a+30),1)                         #Target for 31th row
$tar312 = $ws.Cells.Item(($a+30),2)
$tar313 = $ws.Cells.Item(($a+30),3)
$tar314 = $ws.Cells.Item(($a+30),4)
$tar315 = $ws.Cells.Item(($a+30),5)
$tar316 = $ws.Cells.Item(($a+30),6)
$tar317 = $ws.Cells.Item(($a+30),7)                         
$tar31last = $ws.Cells.Item(($a+30),$b)

$tar321 = $ws.Cells.Item(($a+31),1)                         #Target for 32th row
$tar322 = $ws.Cells.Item(($a+31),2)
$tar323 = $ws.Cells.Item(($a+31),3)
$tar324 = $ws.Cells.Item(($a+31),4)
$tar325 = $ws.Cells.Item(($a+31),5)
$tar326 = $ws.Cells.Item(($a+31),6)
$tar327 = $ws.Cells.Item(($a+31),7)                         
$tar32last = $ws.Cells.Item(($a+31),$b)

$tar331 = $ws.Cells.Item(($a+32),1)                         #Target for 33th row
$tar332 = $ws.Cells.Item(($a+32),2)
$tar333 = $ws.Cells.Item(($a+32),3)
$tar334 = $ws.Cells.Item(($a+32),4)
$tar335 = $ws.Cells.Item(($a+32),5)
$tar336 = $ws.Cells.Item(($a+32),6)
$tar337 = $ws.Cells.Item(($a+32),7)                         
$tar33last = $ws.Cells.Item(($a+32),$b)

$tar341 = $ws.Cells.Item(($a+33),1)                         #Target for 34th row
$tar342 = $ws.Cells.Item(($a+33),2)
$tar343 = $ws.Cells.Item(($a+33),3)
$tar344 = $ws.Cells.Item(($a+33),4)
$tar345 = $ws.Cells.Item(($a+33),5)
$tar346 = $ws.Cells.Item(($a+33),6)
$tar347 = $ws.Cells.Item(($a+33),7)                         
$tar34last = $ws.Cells.Item(($a+33),$b)

$tar351 = $ws.Cells.Item(($a+34),1)                         #Target for 35th row
$tar352 = $ws.Cells.Item(($a+34),2)
$tar353 = $ws.Cells.Item(($a+34),3)
$tar354 = $ws.Cells.Item(($a+34),4)
$tar355 = $ws.Cells.Item(($a+34),5)
$tar356 = $ws.Cells.Item(($a+34),6)
$tar357 = $ws.Cells.Item(($a+34),7)                         
$tar35last = $ws.Cells.Item(($a+34),$b)

$tar361 = $ws.Cells.Item(($a+35),1)                         #Target for 36th row
$tar362 = $ws.Cells.Item(($a+35),2)
$tar363 = $ws.Cells.Item(($a+35),3)
$tar364 = $ws.Cells.Item(($a+35),4)
$tar365 = $ws.Cells.Item(($a+35),5)
$tar366 = $ws.Cells.Item(($a+35),6)
$tar367 = $ws.Cells.Item(($a+35),7)                         
$tar36last = $ws.Cells.Item(($a+35),$b)

$tar371 = $ws.Cells.Item(($a+36),1)                         #Target for 37th row
$tar372 = $ws.Cells.Item(($a+36),2)
$tar373 = $ws.Cells.Item(($a+36),3)
$tar374 = $ws.Cells.Item(($a+36),4)
$tar375 = $ws.Cells.Item(($a+36),5)
$tar376 = $ws.Cells.Item(($a+36),6)
$tar377 = $ws.Cells.Item(($a+36),7)                         
$tar37last = $ws.Cells.Item(($a+36),$b)

$tar381 = $ws.Cells.Item(($a+37),1)                         #Target for 38th row
$tar382 = $ws.Cells.Item(($a+37),2)
$tar383 = $ws.Cells.Item(($a+37),3)
$tar384 = $ws.Cells.Item(($a+37),4)
$tar385 = $ws.Cells.Item(($a+37),5)
$tar386 = $ws.Cells.Item(($a+37),6)
$tar387 = $ws.Cells.Item(($a+37),7)                         
$tar38last = $ws.Cells.Item(($a+37),$b)

$tar391 = $ws.Cells.Item(($a+38),1)                         #Target for 39th row
$tar392 = $ws.Cells.Item(($a+38),2)
$tar393 = $ws.Cells.Item(($a+38),3)
$tar394 = $ws.Cells.Item(($a+38),4)
$tar395 = $ws.Cells.Item(($a+38),5)
$tar396 = $ws.Cells.Item(($a+38),6)
$tar397 = $ws.Cells.Item(($a+38),7)                         
$tar39last = $ws.Cells.Item(($a+38),$b)

$tar401 = $ws.Cells.Item(($a+39),1)                         #Target for 40th row
$tar402 = $ws.Cells.Item(($a+39),2)
$tar403 = $ws.Cells.Item(($a+39),3)
$tar404 = $ws.Cells.Item(($a+39),4)
$tar405 = $ws.Cells.Item(($a+39),5)
$tar406 = $ws.Cells.Item(($a+39),6)
$tar407 = $ws.Cells.Item(($a+39),7)                         
$tar40last = $ws.Cells.Item(($a+39),$b)

$tar411 = $ws.Cells.Item(($a+40),1)                         #Target for 41th row
$tar412 = $ws.Cells.Item(($a+40),2)
$tar413 = $ws.Cells.Item(($a+40),3)
$tar414 = $ws.Cells.Item(($a+40),4)
$tar415 = $ws.Cells.Item(($a+40),5)
$tar416 = $ws.Cells.Item(($a+40),6)
$tar417 = $ws.Cells.Item(($a+40),7)                         
$tar41last = $ws.Cells.Item(($a+40),$b)

$tar421 = $ws.Cells.Item(($a+41),1)                         #Target for 42th row
$tar422 = $ws.Cells.Item(($a+41),2)
$tar423 = $ws.Cells.Item(($a+41),3)
$tar424 = $ws.Cells.Item(($a+41),4)
$tar425 = $ws.Cells.Item(($a+41),5)
$tar426 = $ws.Cells.Item(($a+41),6)
$tar427 = $ws.Cells.Item(($a+41),7)                         
$tar42last = $ws.Cells.Item(($a+41),$b)

$tar431 = $ws.Cells.Item(($a+42),1)                         #Target for 43th row
$tar432 = $ws.Cells.Item(($a+42),2)
$tar433 = $ws.Cells.Item(($a+42),3)
$tar434 = $ws.Cells.Item(($a+42),4)
$tar435 = $ws.Cells.Item(($a+42),5)
$tar436 = $ws.Cells.Item(($a+42),6)
$tar437 = $ws.Cells.Item(($a+42),7)                         
$tar43last = $ws.Cells.Item(($a+42),$b)

$tar441 = $ws.Cells.Item(($a+43),1)                         #Target for 44th row
$tar442 = $ws.Cells.Item(($a+43),2)
$tar443 = $ws.Cells.Item(($a+43),3)
$tar444 = $ws.Cells.Item(($a+43),4)
$tar445 = $ws.Cells.Item(($a+43),5)
$tar446 = $ws.Cells.Item(($a+43),6)
$tar447 = $ws.Cells.Item(($a+43),7)                         
$tar44last = $ws.Cells.Item(($a+43),$b)

$tar451 = $ws.Cells.Item(($a+44),1)                         #Target for 45th row
$tar452 = $ws.Cells.Item(($a+44),2)
$tar453 = $ws.Cells.Item(($a+44),3)
$tar454 = $ws.Cells.Item(($a+44),4)
$tar455 = $ws.Cells.Item(($a+44),5)
$tar456 = $ws.Cells.Item(($a+44),6)
$tar457 = $ws.Cells.Item(($a+44),7)                         
$tar45last = $ws.Cells.Item(($a+44),$b)

$tar461 = $ws.Cells.Item(($a+45),1)                         #Target for 46th row
$tar462 = $ws.Cells.Item(($a+45),2)
$tar463 = $ws.Cells.Item(($a+45),3)
$tar464 = $ws.Cells.Item(($a+45),4)
$tar465 = $ws.Cells.Item(($a+45),5)
$tar466 = $ws.Cells.Item(($a+45),6)
$tar467 = $ws.Cells.Item(($a+45),7)                         
$tar46last = $ws.Cells.Item(($a+45),$b)

$tar471 = $ws.Cells.Item(($a+46),1)                         #Target for 47th row
$tar472 = $ws.Cells.Item(($a+46),2)
$tar473 = $ws.Cells.Item(($a+46),3)
$tar474 = $ws.Cells.Item(($a+46),4)
$tar475 = $ws.Cells.Item(($a+46),5)
$tar476 = $ws.Cells.Item(($a+46),6)
$tar477 = $ws.Cells.Item(($a+46),7)                         
$tar47last = $ws.Cells.Item(($a+46),$b)

$tar481 = $ws.Cells.Item(($a+47),1)                         #Target for 48th row
$tar482 = $ws.Cells.Item(($a+47),2)
$tar483 = $ws.Cells.Item(($a+47),3)
$tar484 = $ws.Cells.Item(($a+47),4)
$tar485 = $ws.Cells.Item(($a+47),5)
$tar486 = $ws.Cells.Item(($a+47),6)
$tar487 = $ws.Cells.Item(($a+47),7)                         
$tar48last = $ws.Cells.Item(($a+47),$b)

$tar491 = $ws.Cells.Item(($a+48),1)                         #Target for 49th row
$tar492 = $ws.Cells.Item(($a+48),2)
$tar493 = $ws.Cells.Item(($a+48),3)
$tar494 = $ws.Cells.Item(($a+48),4)
$tar495 = $ws.Cells.Item(($a+48),5)
$tar496 = $ws.Cells.Item(($a+48),6)
$tar497 = $ws.Cells.Item(($a+48),7)                         
$tar49last = $ws.Cells.Item(($a+48),$b)

$tar501 = $ws.Cells.Item(($a+49),1)                         #Target for 50th row
$tar502 = $ws.Cells.Item(($a+49),2)
$tar503 = $ws.Cells.Item(($a+49),3)
$tar504 = $ws.Cells.Item(($a+49),4)
$tar505 = $ws.Cells.Item(($a+49),5)
$tar506 = $ws.Cells.Item(($a+49),6)
$tar507 = $ws.Cells.Item(($a+49),7)                         
$tar50last = $ws.Cells.Item(($a+49),$b)

$tar511 = $ws.Cells.Item(($a+50),1)                         #Target for 51th row
$tar512 = $ws.Cells.Item(($a+50),2)
$tar513 = $ws.Cells.Item(($a+50),3)
$tar514 = $ws.Cells.Item(($a+50),4)
$tar515 = $ws.Cells.Item(($a+50),5)
$tar516 = $ws.Cells.Item(($a+50),6)
$tar517 = $ws.Cells.Item(($a+50),7)                         
$tar51last = $ws.Cells.Item(($a+50),$b)

$tar521 = $ws.Cells.Item(($a+51),1)                         #Target for 52th row
$tar522 = $ws.Cells.Item(($a+51),2)
$tar523 = $ws.Cells.Item(($a+51),3)
$tar524 = $ws.Cells.Item(($a+51),4)
$tar525 = $ws.Cells.Item(($a+51),5)
$tar526 = $ws.Cells.Item(($a+51),6)
$tar527 = $ws.Cells.Item(($a+51),7)                         
$tar52last = $ws.Cells.Item(($a+51),$b)

$tar531 = $ws.Cells.Item(($a+52),1)                         #Target for 53th row
$tar532 = $ws.Cells.Item(($a+52),2)
$tar533 = $ws.Cells.Item(($a+52),3)
$tar534 = $ws.Cells.Item(($a+52),4)
$tar535 = $ws.Cells.Item(($a+52),5)
$tar536 = $ws.Cells.Item(($a+52),6)
$tar537 = $ws.Cells.Item(($a+52),7)                         
$tar53last = $ws.Cells.Item(($a+52),$b)

$tar541 = $ws.Cells.Item(($a+53),1)                         #Target for 54th row
$tar542 = $ws.Cells.Item(($a+53),2)
$tar543 = $ws.Cells.Item(($a+53),3)
$tar544 = $ws.Cells.Item(($a+53),4)
$tar545 = $ws.Cells.Item(($a+53),5)
$tar546 = $ws.Cells.Item(($a+53),6)
$tar547 = $ws.Cells.Item(($a+53),7)                         
$tar54last = $ws.Cells.Item(($a+53),$b)

$tar551 = $ws.Cells.Item(($a+54),1)                         #Target for 55th row
$tar552 = $ws.Cells.Item(($a+54),2)
$tar553 = $ws.Cells.Item(($a+54),3)
$tar554 = $ws.Cells.Item(($a+54),4)
$tar555 = $ws.Cells.Item(($a+54),5)
$tar556 = $ws.Cells.Item(($a+54),6)
$tar557 = $ws.Cells.Item(($a+54),7)                         
$tar55last = $ws.Cells.Item(($a+54),$b)

$tar561 = $ws.Cells.Item(($a+55),1)                         #Target for 56th row
$tar562 = $ws.Cells.Item(($a+55),2)
$tar563 = $ws.Cells.Item(($a+55),3)
$tar564 = $ws.Cells.Item(($a+55),4)
$tar565 = $ws.Cells.Item(($a+55),5)
$tar566 = $ws.Cells.Item(($a+55),6)
$tar567 = $ws.Cells.Item(($a+55),7)                         
$tar56last = $ws.Cells.Item(($a+55),$b)

$tar571 = $ws.Cells.Item(($a+56),1)                         #Target for 57th row
$tar572 = $ws.Cells.Item(($a+56),2)
$tar573 = $ws.Cells.Item(($a+56),3)
$tar574 = $ws.Cells.Item(($a+56),4)
$tar575 = $ws.Cells.Item(($a+56),5)
$tar576 = $ws.Cells.Item(($a+56),6)
$tar577 = $ws.Cells.Item(($a+56),7)                         
$tar57last = $ws.Cells.Item(($a+56),$b)

$tar581 = $ws.Cells.Item(($a+57),1)                         #Target for 58th row
$tar582 = $ws.Cells.Item(($a+57),2)
$tar583 = $ws.Cells.Item(($a+57),3)
$tar584 = $ws.Cells.Item(($a+57),4)
$tar585 = $ws.Cells.Item(($a+57),5)
$tar586 = $ws.Cells.Item(($a+57),6)
$tar587 = $ws.Cells.Item(($a+57),7)                         
$tar58last = $ws.Cells.Item(($a+57),$b)

$tar591 = $ws.Cells.Item(($a+58),1)                         #Target for 59th row
$tar592 = $ws.Cells.Item(($a+58),2)
$tar593 = $ws.Cells.Item(($a+58),3)
$tar594 = $ws.Cells.Item(($a+58),4)
$tar595 = $ws.Cells.Item(($a+58),5)
$tar596 = $ws.Cells.Item(($a+58),6)
$tar597 = $ws.Cells.Item(($a+58),7)                         
$tar59last = $ws.Cells.Item(($a+58),$b)

$tar601 = $ws.Cells.Item(($a+59),1)                         #Target for 60th row
$tar602 = $ws.Cells.Item(($a+59),2)
$tar603 = $ws.Cells.Item(($a+59),3)
$tar604 = $ws.Cells.Item(($a+59),4)
$tar605 = $ws.Cells.Item(($a+59),5)
$tar606 = $ws.Cells.Item(($a+59),6)
$tar607 = $ws.Cells.Item(($a+59),7)                         
$tar60last = $ws.Cells.Item(($a+59),$b)

$tar611 = $ws.Cells.Item(($a+60),1)                         #Target for 61th row
$tar612 = $ws.Cells.Item(($a+60),2)
$tar613 = $ws.Cells.Item(($a+60),3)
$tar614 = $ws.Cells.Item(($a+60),4)
$tar615 = $ws.Cells.Item(($a+60),5)
$tar616 = $ws.Cells.Item(($a+60),6)
$tar617 = $ws.Cells.Item(($a+60),7)                         
$tar61last = $ws.Cells.Item(($a+60),$b)

$tar621 = $ws.Cells.Item(($a+61),1)                         #Target for 62th row
$tar622 = $ws.Cells.Item(($a+61),2)
$tar623 = $ws.Cells.Item(($a+61),3)
$tar624 = $ws.Cells.Item(($a+61),4)
$tar625 = $ws.Cells.Item(($a+61),5)
$tar626 = $ws.Cells.Item(($a+61),6)
$tar627 = $ws.Cells.Item(($a+61),7)                         
$tar62last = $ws.Cells.Item(($a+61),$b)

$tar631 = $ws.Cells.Item(($a+62),1)                         #Target for 63th row
$tar632 = $ws.Cells.Item(($a+62),2)
$tar633 = $ws.Cells.Item(($a+62),3)
$tar634 = $ws.Cells.Item(($a+62),4)
$tar635 = $ws.Cells.Item(($a+62),5)
$tar636 = $ws.Cells.Item(($a+62),6)
$tar637 = $ws.Cells.Item(($a+62),7)                         
$tar63last = $ws.Cells.Item(($a+62),$b)

$tar641 = $ws.Cells.Item(($a+63),1)                         #Target for 64th row
$tar642 = $ws.Cells.Item(($a+63),2)
$tar643 = $ws.Cells.Item(($a+63),3)
$tar644 = $ws.Cells.Item(($a+63),4)
$tar645 = $ws.Cells.Item(($a+63),5)
$tar646 = $ws.Cells.Item(($a+63),6)
$tar647 = $ws.Cells.Item(($a+63),7)                         
$tar64last = $ws.Cells.Item(($a+63),$b)

$tar651 = $ws.Cells.Item(($a+64),1)                         #Target for 65th row
$tar652 = $ws.Cells.Item(($a+64),2)
$tar653 = $ws.Cells.Item(($a+64),3)
$tar654 = $ws.Cells.Item(($a+64),4)
$tar655 = $ws.Cells.Item(($a+64),5)
$tar656 = $ws.Cells.Item(($a+64),6)
$tar657 = $ws.Cells.Item(($a+64),7)                         
$tar65last = $ws.Cells.Item(($a+64),$b)

$tar661 = $ws.Cells.Item(($a+65),1)                         #Target for 66th row
$tar662 = $ws.Cells.Item(($a+65),2)
$tar663 = $ws.Cells.Item(($a+65),3)
$tar664 = $ws.Cells.Item(($a+65),4)
$tar665 = $ws.Cells.Item(($a+65),5)
$tar666 = $ws.Cells.Item(($a+65),6)
$tar667 = $ws.Cells.Item(($a+65),7)                         
$tar66last = $ws.Cells.Item(($a+65),$b)

$tar671 = $ws.Cells.Item(($a+66),1)                         #Target for 67th row
$tar672 = $ws.Cells.Item(($a+66),2)
$tar673 = $ws.Cells.Item(($a+66),3)
$tar674 = $ws.Cells.Item(($a+66),4)
$tar675 = $ws.Cells.Item(($a+66),5)
$tar676 = $ws.Cells.Item(($a+66),6)
$tar677 = $ws.Cells.Item(($a+66),7)                         
$tar67last = $ws.Cells.Item(($a+66),$b)

$tar681 = $ws.Cells.Item(($a+67),1)                         #Target for 68th row
$tar682 = $ws.Cells.Item(($a+67),2)
$tar683 = $ws.Cells.Item(($a+67),3)
$tar684 = $ws.Cells.Item(($a+67),4)
$tar685 = $ws.Cells.Item(($a+67),5)
$tar686 = $ws.Cells.Item(($a+67),6)
$tar687 = $ws.Cells.Item(($a+67),7)                         
$tar68last = $ws.Cells.Item(($a+67),$b)

$tar691 = $ws.Cells.Item(($a+68),1)                         #Target for 69th row
$tar692 = $ws.Cells.Item(($a+68),2)
$tar693 = $ws.Cells.Item(($a+68),3)
$tar694 = $ws.Cells.Item(($a+68),4)
$tar695 = $ws.Cells.Item(($a+68),5)
$tar696 = $ws.Cells.Item(($a+68),6)
$tar697 = $ws.Cells.Item(($a+68),7)                         
$tar69last = $ws.Cells.Item(($a+68),$b)

$tar701 = $ws.Cells.Item(($a+69),1)                         #Target for 70th row
$tar702 = $ws.Cells.Item(($a+69),2)
$tar703 = $ws.Cells.Item(($a+69),3)
$tar704 = $ws.Cells.Item(($a+69),4)
$tar705 = $ws.Cells.Item(($a+69),5)
$tar706 = $ws.Cells.Item(($a+69),6)
$tar707 = $ws.Cells.Item(($a+69),7)                         
$tar70last = $ws.Cells.Item(($a+69),$b)

$tar711 = $ws.Cells.Item(($a+70),1)                         #Target for 71th row
$tar712 = $ws.Cells.Item(($a+70),2)
$tar713 = $ws.Cells.Item(($a+70),3)
$tar714 = $ws.Cells.Item(($a+70),4)
$tar715 = $ws.Cells.Item(($a+70),5)
$tar716 = $ws.Cells.Item(($a+70),6)
$tar717 = $ws.Cells.Item(($a+70),7)                         
$tar71last = $ws.Cells.Item(($a+70),$b)

$tar721 = $ws.Cells.Item(($a+71),1)                         #Target for 72th row
$tar722 = $ws.Cells.Item(($a+71),2)
$tar723 = $ws.Cells.Item(($a+71),3)
$tar724 = $ws.Cells.Item(($a+71),4)
$tar725 = $ws.Cells.Item(($a+71),5)
$tar726 = $ws.Cells.Item(($a+71),6)
$tar727 = $ws.Cells.Item(($a+71),7)                         
$tar72last = $ws.Cells.Item(($a+71),$b)

$tar731 = $ws.Cells.Item(($a+72),1)                         #Target for 73th row
$tar732 = $ws.Cells.Item(($a+72),2)
$tar733 = $ws.Cells.Item(($a+72),3)
$tar734 = $ws.Cells.Item(($a+72),4)
$tar735 = $ws.Cells.Item(($a+72),5)
$tar736 = $ws.Cells.Item(($a+72),6)
$tar737 = $ws.Cells.Item(($a+72),7)                         
$tar73last = $ws.Cells.Item(($a+72),$b)

$tar741 = $ws.Cells.Item(($a+73),1)                         #Target for 74th row
$tar742 = $ws.Cells.Item(($a+73),2)
$tar743 = $ws.Cells.Item(($a+73),3)
$tar744 = $ws.Cells.Item(($a+73),4)
$tar745 = $ws.Cells.Item(($a+73),5)
$tar746 = $ws.Cells.Item(($a+73),6)
$tar747 = $ws.Cells.Item(($a+73),7)                         
$tar74last = $ws.Cells.Item(($a+73),$b)

$tar751 = $ws.Cells.Item(($a+74),1)                         #Target for 75th row
$tar752 = $ws.Cells.Item(($a+74),2)
$tar753 = $ws.Cells.Item(($a+74),3)
$tar754 = $ws.Cells.Item(($a+74),4)
$tar755 = $ws.Cells.Item(($a+74),5)
$tar756 = $ws.Cells.Item(($a+74),6)
$tar757 = $ws.Cells.Item(($a+74),7)                         
$tar75last = $ws.Cells.Item(($a+74),$b)

$tar761 = $ws.Cells.Item(($a+75),1)                         #Target for 76th row
$tar762 = $ws.Cells.Item(($a+75),2)
$tar763 = $ws.Cells.Item(($a+75),3)
$tar764 = $ws.Cells.Item(($a+75),4)
$tar765 = $ws.Cells.Item(($a+75),5)
$tar766 = $ws.Cells.Item(($a+75),6)
$tar767 = $ws.Cells.Item(($a+75),7)                         
$tar76last = $ws.Cells.Item(($a+75),$b)

$tar771 = $ws.Cells.Item(($a+76),1)                         #Target for 77th row
$tar772 = $ws.Cells.Item(($a+76),2)
$tar773 = $ws.Cells.Item(($a+76),3)
$tar774 = $ws.Cells.Item(($a+76),4)
$tar775 = $ws.Cells.Item(($a+76),5)
$tar776 = $ws.Cells.Item(($a+76),6)
$tar777 = $ws.Cells.Item(($a+76),7)                         
$tar77last = $ws.Cells.Item(($a+76),$b)

$tar781 = $ws.Cells.Item(($a+77),1)                         #Target for 78th row
$tar782 = $ws.Cells.Item(($a+77),2)
$tar783 = $ws.Cells.Item(($a+77),3)
$tar784 = $ws.Cells.Item(($a+77),4)
$tar785 = $ws.Cells.Item(($a+77),5)
$tar786 = $ws.Cells.Item(($a+77),6)
$tar787 = $ws.Cells.Item(($a+77),7)                         
$tar78last = $ws.Cells.Item(($a+77),$b)

$tar791 = $ws.Cells.Item(($a+78),1)                         #Target for 79th row
$tar792 = $ws.Cells.Item(($a+78),2)
$tar793 = $ws.Cells.Item(($a+78),3)
$tar794 = $ws.Cells.Item(($a+78),4)
$tar795 = $ws.Cells.Item(($a+78),5)
$tar796 = $ws.Cells.Item(($a+78),6)
$tar797 = $ws.Cells.Item(($a+78),7)                         
$tar79last = $ws.Cells.Item(($a+78),$b)

$tar801 = $ws.Cells.Item(($a+79),1)                         #Target for 80th row
$tar802 = $ws.Cells.Item(($a+79),2)
$tar803 = $ws.Cells.Item(($a+79),3)
$tar804 = $ws.Cells.Item(($a+79),4)
$tar805 = $ws.Cells.Item(($a+79),5)
$tar806 = $ws.Cells.Item(($a+79),6)
$tar807 = $ws.Cells.Item(($a+79),7)                         
$tar80last = $ws.Cells.Item(($a+79),$b)

$tar811 = $ws.Cells.Item(($a+80),1)                         #Target for 81th row
$tar812 = $ws.Cells.Item(($a+80),2)
$tar813 = $ws.Cells.Item(($a+80),3)
$tar814 = $ws.Cells.Item(($a+80),4)
$tar815 = $ws.Cells.Item(($a+80),5)
$tar816 = $ws.Cells.Item(($a+80),6)
$tar817 = $ws.Cells.Item(($a+80),7)                         
$tar81last = $ws.Cells.Item(($a+80),$b)

$tar821 = $ws.Cells.Item(($a+81),1)                         #Target for 82th row
$tar822 = $ws.Cells.Item(($a+81),2)
$tar823 = $ws.Cells.Item(($a+81),3)
$tar824 = $ws.Cells.Item(($a+81),4)
$tar825 = $ws.Cells.Item(($a+81),5)
$tar826 = $ws.Cells.Item(($a+81),6)
$tar827 = $ws.Cells.Item(($a+81),7)                         
$tar82last = $ws.Cells.Item(($a+81),$b)

$tar831 = $ws.Cells.Item(($a+82),1)                         #Target for 83th row
$tar832 = $ws.Cells.Item(($a+82),2)
$tar833 = $ws.Cells.Item(($a+82),3)
$tar834 = $ws.Cells.Item(($a+82),4)
$tar835 = $ws.Cells.Item(($a+82),5)
$tar836 = $ws.Cells.Item(($a+82),6)
$tar837 = $ws.Cells.Item(($a+82),7)                         
$tar83last = $ws.Cells.Item(($a+82),$b)

$tar841 = $ws.Cells.Item(($a+83),1)                         #Target for 84th row
$tar842 = $ws.Cells.Item(($a+83),2)
$tar843 = $ws.Cells.Item(($a+83),3)
$tar844 = $ws.Cells.Item(($a+83),4)
$tar845 = $ws.Cells.Item(($a+83),5)
$tar846 = $ws.Cells.Item(($a+83),6)
$tar847 = $ws.Cells.Item(($a+83),7)                         
$tar84last = $ws.Cells.Item(($a+83),$b)

$tar851 = $ws.Cells.Item(($a+84),1)                         #Target for 85th row
$tar852 = $ws.Cells.Item(($a+84),2)
$tar853 = $ws.Cells.Item(($a+84),3)
$tar854 = $ws.Cells.Item(($a+84),4)
$tar855 = $ws.Cells.Item(($a+84),5)
$tar856 = $ws.Cells.Item(($a+84),6)
$tar857 = $ws.Cells.Item(($a+84),7)                         
$tar85last = $ws.Cells.Item(($a+84),$b)

$tar861 = $ws.Cells.Item(($a+85),1)                         #Target for 86th row
$tar862 = $ws.Cells.Item(($a+85),2)
$tar863 = $ws.Cells.Item(($a+85),3)
$tar864 = $ws.Cells.Item(($a+85),4)
$tar865 = $ws.Cells.Item(($a+85),5)
$tar866 = $ws.Cells.Item(($a+85),6)
$tar867 = $ws.Cells.Item(($a+85),7)                         
$tar86last = $ws.Cells.Item(($a+85),$b)

$tar871 = $ws.Cells.Item(($a+86),1)                         #Target for 87th row
$tar872 = $ws.Cells.Item(($a+86),2)
$tar873 = $ws.Cells.Item(($a+86),3)
$tar874 = $ws.Cells.Item(($a+86),4)
$tar875 = $ws.Cells.Item(($a+86),5)
$tar876 = $ws.Cells.Item(($a+86),6)
$tar877 = $ws.Cells.Item(($a+86),7)                         
$tar87last = $ws.Cells.Item(($a+86),$b)

$tar881 = $ws.Cells.Item(($a+87),1)                         #Target for 88th row
$tar882 = $ws.Cells.Item(($a+87),2)
$tar883 = $ws.Cells.Item(($a+87),3)
$tar884 = $ws.Cells.Item(($a+87),4)
$tar885 = $ws.Cells.Item(($a+87),5)
$tar886 = $ws.Cells.Item(($a+87),6)
$tar887 = $ws.Cells.Item(($a+87),7)                         
$tar88last = $ws.Cells.Item(($a+87),$b)

$tar891 = $ws.Cells.Item(($a+88),1)                         #Target for 89th row
$tar892 = $ws.Cells.Item(($a+88),2)
$tar893 = $ws.Cells.Item(($a+88),3)
$tar894 = $ws.Cells.Item(($a+88),4)
$tar895 = $ws.Cells.Item(($a+88),5)
$tar896 = $ws.Cells.Item(($a+88),6)
$tar897 = $ws.Cells.Item(($a+88),7)                         
$tar89last = $ws.Cells.Item(($a+88),$b)

$tar901 = $ws.Cells.Item(($a+89),1)                         #Target for 90th row
$tar902 = $ws.Cells.Item(($a+89),2)
$tar903 = $ws.Cells.Item(($a+89),3)
$tar904 = $ws.Cells.Item(($a+89),4)
$tar905 = $ws.Cells.Item(($a+89),5)
$tar906 = $ws.Cells.Item(($a+89),6)
$tar907 = $ws.Cells.Item(($a+89),7)                         
$tar90last = $ws.Cells.Item(($a+89),$b)

$tar911 = $ws.Cells.Item(($a+90),1)                         #Target for 91th row
$tar912 = $ws.Cells.Item(($a+90),2)
$tar913 = $ws.Cells.Item(($a+90),3)
$tar914 = $ws.Cells.Item(($a+90),4)
$tar915 = $ws.Cells.Item(($a+90),5)
$tar916 = $ws.Cells.Item(($a+90),6)
$tar917 = $ws.Cells.Item(($a+90),7)                         
$tar91last = $ws.Cells.Item(($a+90),$b)

$tar921 = $ws.Cells.Item(($a+91),1)                         #Target for 92th row
$tar922 = $ws.Cells.Item(($a+91),2)
$tar923 = $ws.Cells.Item(($a+91),3)
$tar924 = $ws.Cells.Item(($a+91),4)
$tar925 = $ws.Cells.Item(($a+91),5)
$tar926 = $ws.Cells.Item(($a+91),6)
$tar927 = $ws.Cells.Item(($a+91),7)                         
$tar92last = $ws.Cells.Item(($a+91),$b)

$tar931 = $ws.Cells.Item(($a+92),1)                         #Target for 93th row
$tar932 = $ws.Cells.Item(($a+92),2)
$tar933 = $ws.Cells.Item(($a+92),3)
$tar934 = $ws.Cells.Item(($a+92),4)
$tar935 = $ws.Cells.Item(($a+92),5)
$tar936 = $ws.Cells.Item(($a+92),6)
$tar937 = $ws.Cells.Item(($a+92),7)                         
$tar93last = $ws.Cells.Item(($a+92),$b)

$tar941 = $ws.Cells.Item(($a+93),1)                         #Target for 94th row
$tar942 = $ws.Cells.Item(($a+93),2)
$tar943 = $ws.Cells.Item(($a+93),3)
$tar944 = $ws.Cells.Item(($a+93),4)
$tar945 = $ws.Cells.Item(($a+93),5)
$tar946 = $ws.Cells.Item(($a+93),6)
$tar947 = $ws.Cells.Item(($a+93),7)                         
$tar94last = $ws.Cells.Item(($a+93),$b)

$tar951 = $ws.Cells.Item(($a+94),1)                         #Target for 95th row
$tar952 = $ws.Cells.Item(($a+94),2)
$tar953 = $ws.Cells.Item(($a+94),3)
$tar954 = $ws.Cells.Item(($a+94),4)
$tar955 = $ws.Cells.Item(($a+94),5)
$tar956 = $ws.Cells.Item(($a+94),6)
$tar957 = $ws.Cells.Item(($a+94),7)                         
$tar95last = $ws.Cells.Item(($a+94),$b)

$tar961 = $ws.Cells.Item(($a+95),1)                         #Target for 96th row
$tar962 = $ws.Cells.Item(($a+95),2)
$tar963 = $ws.Cells.Item(($a+95),3)
$tar964 = $ws.Cells.Item(($a+95),4)
$tar965 = $ws.Cells.Item(($a+95),5)
$tar966 = $ws.Cells.Item(($a+95),6)
$tar967 = $ws.Cells.Item(($a+95),7)                         
$tar96last = $ws.Cells.Item(($a+95),$b)

$tar971 = $ws.Cells.Item(($a+96),1)                         #Target for 97th row
$tar972 = $ws.Cells.Item(($a+96),2)
$tar973 = $ws.Cells.Item(($a+96),3)
$tar974 = $ws.Cells.Item(($a+96),4)
$tar975 = $ws.Cells.Item(($a+96),5)
$tar976 = $ws.Cells.Item(($a+96),6)
$tar977 = $ws.Cells.Item(($a+96),7)                         
$tar97last = $ws.Cells.Item(($a+96),$b)

$tar981 = $ws.Cells.Item(($a+97),1)                         #Target for 98th row
$tar982 = $ws.Cells.Item(($a+97),2)
$tar983 = $ws.Cells.Item(($a+97),3)
$tar984 = $ws.Cells.Item(($a+97),4)
$tar985 = $ws.Cells.Item(($a+97),5)
$tar986 = $ws.Cells.Item(($a+97),6)
$tar987 = $ws.Cells.Item(($a+97),7)                         
$tar98last = $ws.Cells.Item(($a+97),$b)

$tar991 = $ws.Cells.Item(($a+98),1)                         #Target for 99th row
$tar992 = $ws.Cells.Item(($a+98),2)
$tar993 = $ws.Cells.Item(($a+98),3)
$tar994 = $ws.Cells.Item(($a+98),4)
$tar995 = $ws.Cells.Item(($a+98),5)
$tar996 = $ws.Cells.Item(($a+98),6)
$tar997 = $ws.Cells.Item(($a+98),7)                         
$tar99last = $ws.Cells.Item(($a+98),$b)

$tar1001 = $ws.Cells.Item(($a+99),1)                         #Target for 100th row
$tar1002 = $ws.Cells.Item(($a+99),2)
$tar1003 = $ws.Cells.Item(($a+99),3)
$tar1004 = $ws.Cells.Item(($a+99),4)
$tar1005 = $ws.Cells.Item(($a+99),5)
$tar1006 = $ws.Cells.Item(($a+99),6)
$tar1007 = $ws.Cells.Item(($a+99),7)                         
$tar100last = $ws.Cells.Item(($a+99),$b)
#------------------------------------------------------------------------------------------
$sor1.Copy($tar1)                                          #Action
$sor2.Copy($tar2)
$sor3.Copy($tar3)
$sor4.Copy($tar4)
$sor5.Copy($tar5)
$sor6.Copy($tar6)
$sor7.Copy($tar7)
$sorlast.Copy($tarlast)

$sor21.Copy($tar21)                                        #Action for second row
$sor22.Copy($tar22)
$sor23.Copy($tar23)
$sor24.Copy($tar24)
$sor25.Copy($tar25)
$sor26.Copy($tar26)
$sor27.Copy($tar27)
$sor2last.Copy($tar2last)

$sor31.Copy($tar31)                                        #Action for 3th row
$sor32.Copy($tar32)
$sor33.Copy($tar33)
$sor34.Copy($tar34)
$sor35.Copy($tar35)
$sor36.Copy($tar36)
$sor37.Copy($tar37)
$sor3last.Copy($tar3last)

$sor41.Copy($tar41)                                        #Action for 4th row
$sor42.Copy($tar42)
$sor43.Copy($tar43)
$sor44.Copy($tar44)
$sor45.Copy($tar45)
$sor46.Copy($tar46)
$sor47.Copy($tar47)
$sor4last.Copy($tar4last)

$sor51.Copy($tar51)                                        #Action for 5th row
$sor52.Copy($tar52)
$sor53.Copy($tar53)
$sor54.Copy($tar54)
$sor55.Copy($tar55)
$sor56.Copy($tar56)
$sor57.Copy($tar57)
$sor5last.Copy($tar5last)

$sor61.Copy($tar61)                                        #Action for 6th row
$sor62.Copy($tar62)
$sor63.Copy($tar63)
$sor64.Copy($tar64)
$sor65.Copy($tar65)
$sor66.Copy($tar66)
$sor67.Copy($tar67)
$sor6last.Copy($tar6last)

$sor71.Copy($tar71)                                        #Action for 7th row
$sor72.Copy($tar72)
$sor73.Copy($tar73)
$sor74.Copy($tar74)
$sor75.Copy($tar75)
$sor76.Copy($tar76)
$sor77.Copy($tar77)
$sor7last.Copy($tar7last)

$sor81.Copy($tar81)                                        #Action for 8th row
$sor82.Copy($tar82)
$sor83.Copy($tar83)
$sor84.Copy($tar84)
$sor85.Copy($tar85)
$sor86.Copy($tar86)
$sor87.Copy($tar87)
$sor8last.Copy($tar8last)

$sor91.Copy($tar91)                                        #Action for 9th row
$sor92.Copy($tar92)
$sor93.Copy($tar93)
$sor94.Copy($tar94)
$sor95.Copy($tar95)
$sor96.Copy($tar96)
$sor97.Copy($tar97)
$sor9last.Copy($tar9last)

$sor101.Copy($tar101)                                        #Action for 10th row
$sor102.Copy($tar102)
$sor103.Copy($tar103)
$sor104.Copy($tar104)
$sor105.Copy($tar105)
$sor106.Copy($tar106)
$sor107.Copy($tar107)
$sor10last.Copy($tar10last)

$sor111.Copy($tar111)                                        #Action for 11th row
$sor112.Copy($tar112)
$sor113.Copy($tar113)
$sor114.Copy($tar114)
$sor115.Copy($tar115)
$sor116.Copy($tar116)
$sor117.Copy($tar117)
$sor11last.Copy($tar11last)

$sor121.Copy($tar121)                                        #Action for 12th row
$sor122.Copy($tar122)
$sor123.Copy($tar123)
$sor124.Copy($tar124)
$sor125.Copy($tar125)
$sor126.Copy($tar126)
$sor127.Copy($tar127)
$sor12last.Copy($tar12last)

$sor131.Copy($tar131)                                        #Action for 13th row
$sor132.Copy($tar132)
$sor133.Copy($tar133)
$sor134.Copy($tar134)
$sor135.Copy($tar135)
$sor136.Copy($tar136)
$sor137.Copy($tar137)
$sor13last.Copy($tar13last)

$sor141.Copy($tar141)                                        #Action for 14th row
$sor142.Copy($tar142)
$sor143.Copy($tar143)
$sor144.Copy($tar144)
$sor145.Copy($tar145)
$sor146.Copy($tar146)
$sor147.Copy($tar147)
$sor14last.Copy($tar14last)

$sor151.Copy($tar151)                                        #Action for 15th row
$sor152.Copy($tar152)
$sor153.Copy($tar153)
$sor154.Copy($tar154)
$sor155.Copy($tar155)
$sor156.Copy($tar156)
$sor157.Copy($tar157)
$sor15last.Copy($tar15last)

$sor161.Copy($tar161)                                        #Action for 16th row
$sor162.Copy($tar162)
$sor163.Copy($tar163)
$sor164.Copy($tar164)
$sor165.Copy($tar165)
$sor166.Copy($tar166)
$sor167.Copy($tar167)
$sor16last.Copy($tar16last)

$sor171.Copy($tar171)                                        #Action for 17th row
$sor172.Copy($tar172)
$sor173.Copy($tar173)
$sor174.Copy($tar174)
$sor175.Copy($tar175)
$sor176.Copy($tar176)
$sor177.Copy($tar177)
$sor17last.Copy($tar17last)

$sor181.Copy($tar181)                                        #Action for 18th row
$sor182.Copy($tar182)
$sor183.Copy($tar183)
$sor184.Copy($tar184)
$sor185.Copy($tar185)
$sor186.Copy($tar186)
$sor187.Copy($tar187)
$sor18last.Copy($tar18last)

$sor191.Copy($tar191)                                        #Action for 19th row
$sor192.Copy($tar192)
$sor193.Copy($tar193)
$sor194.Copy($tar194)
$sor195.Copy($tar195)
$sor196.Copy($tar196)
$sor197.Copy($tar197)
$sor19last.Copy($tar19last)

$sor201.Copy($tar201)                                        #Action for 20th row
$sor202.Copy($tar202)
$sor203.Copy($tar203)
$sor204.Copy($tar204)
$sor205.Copy($tar205)
$sor206.Copy($tar206)
$sor207.Copy($tar207)
$sor20last.Copy($tar20last)

$sor211.Copy($tar211)                                        #Action for 21th row
$sor212.Copy($tar212)
$sor213.Copy($tar213)
$sor214.Copy($tar214)
$sor215.Copy($tar215)
$sor216.Copy($tar216)
$sor217.Copy($tar217)
$sor21last.Copy($tar21last)

$sor221.Copy($tar221)                                        #Action for 22th row
$sor222.Copy($tar222)
$sor223.Copy($tar223)
$sor224.Copy($tar224)
$sor225.Copy($tar225)
$sor226.Copy($tar226)
$sor227.Copy($tar227)
$sor22last.Copy($tar22last)

$sor231.Copy($tar231)                                        #Action for 23th row
$sor232.Copy($tar232)
$sor233.Copy($tar233)
$sor234.Copy($tar234)
$sor235.Copy($tar235)
$sor236.Copy($tar236)
$sor237.Copy($tar237)
$sor23last.Copy($tar23last)

$sor241.Copy($tar241)                                        #Action for 24th row
$sor242.Copy($tar242)
$sor243.Copy($tar243)
$sor244.Copy($tar244)
$sor245.Copy($tar245)
$sor246.Copy($tar246)
$sor247.Copy($tar247)
$sor24last.Copy($tar24last)

$sor251.Copy($tar251)                                        #Action for 25th row
$sor252.Copy($tar252)
$sor253.Copy($tar253)
$sor254.Copy($tar254)
$sor255.Copy($tar255)
$sor256.Copy($tar256)
$sor257.Copy($tar257)
$sor25last.Copy($tar25last)

$sor261.Copy($tar261)                                        #Action for 26th row
$sor262.Copy($tar262)
$sor263.Copy($tar263)
$sor264.Copy($tar264)
$sor265.Copy($tar265)
$sor266.Copy($tar266)
$sor267.Copy($tar267)
$sor26last.Copy($tar26last)

$sor271.Copy($tar271)                                        #Action for 27th row
$sor272.Copy($tar272)
$sor273.Copy($tar273)
$sor274.Copy($tar274)
$sor275.Copy($tar275)
$sor276.Copy($tar276)
$sor277.Copy($tar277)
$sor27last.Copy($tar27last)

$sor281.Copy($tar281)                                        #Action for 28th row
$sor282.Copy($tar282)
$sor283.Copy($tar283)
$sor284.Copy($tar284)
$sor285.Copy($tar285)
$sor286.Copy($tar286)
$sor287.Copy($tar287)
$sor28last.Copy($tar28last)

$sor291.Copy($tar291)                                        #Action for 29th row
$sor292.Copy($tar292)
$sor293.Copy($tar293)
$sor294.Copy($tar294)
$sor295.Copy($tar295)
$sor296.Copy($tar296)
$sor297.Copy($tar297)
$sor29last.Copy($tar29last)

$sor301.Copy($tar301)                                        #Action for 30th row
$sor302.Copy($tar302)
$sor303.Copy($tar303)
$sor304.Copy($tar304)
$sor305.Copy($tar305)
$sor306.Copy($tar306)
$sor307.Copy($tar307)
$sor30last.Copy($tar30last)

$sor311.Copy($tar311)                                        #Action for 31th row
$sor312.Copy($tar312)
$sor313.Copy($tar313)
$sor314.Copy($tar314)
$sor315.Copy($tar315)
$sor316.Copy($tar316)
$sor317.Copy($tar317)
$sor31last.Copy($tar31last)

$sor321.Copy($tar321)                                        #Action for 32th row
$sor322.Copy($tar322)
$sor323.Copy($tar323)
$sor324.Copy($tar324)
$sor325.Copy($tar325)
$sor326.Copy($tar326)
$sor327.Copy($tar327)
$sor32last.Copy($tar32last)

$sor331.Copy($tar331)                                        #Action for 33th row
$sor332.Copy($tar332)
$sor333.Copy($tar333)
$sor334.Copy($tar334)
$sor335.Copy($tar335)
$sor336.Copy($tar336)
$sor337.Copy($tar337)
$sor33last.Copy($tar33last)

$sor341.Copy($tar341)                                        #Action for 34th row
$sor342.Copy($tar342)
$sor343.Copy($tar343)
$sor344.Copy($tar344)
$sor345.Copy($tar345)
$sor346.Copy($tar346)
$sor347.Copy($tar347)
$sor34last.Copy($tar34last)

$sor351.Copy($tar351)                                        #Action for 35th row
$sor352.Copy($tar352)
$sor353.Copy($tar353)
$sor354.Copy($tar354)
$sor355.Copy($tar355)
$sor356.Copy($tar356)
$sor357.Copy($tar357)
$sor35last.Copy($tar35last)

$sor361.Copy($tar361)                                        #Action for 36th row
$sor362.Copy($tar362)
$sor363.Copy($tar363)
$sor364.Copy($tar364)
$sor365.Copy($tar365)
$sor366.Copy($tar366)
$sor367.Copy($tar367)
$sor36last.Copy($tar36last)

$sor371.Copy($tar371)                                        #Action for 37th row
$sor372.Copy($tar372)
$sor373.Copy($tar373)
$sor374.Copy($tar374)
$sor375.Copy($tar375)
$sor376.Copy($tar376)
$sor377.Copy($tar377)
$sor37last.Copy($tar37last)

$sor381.Copy($tar381)                                        #Action for 38th row
$sor382.Copy($tar382)
$sor383.Copy($tar383)
$sor384.Copy($tar384)
$sor385.Copy($tar385)
$sor386.Copy($tar386)
$sor387.Copy($tar387)
$sor38last.Copy($tar38last)

$sor391.Copy($tar391)                                        #Action for 39th row
$sor392.Copy($tar392)
$sor393.Copy($tar393)
$sor394.Copy($tar394)
$sor395.Copy($tar395)
$sor396.Copy($tar396)
$sor397.Copy($tar397)
$sor39last.Copy($tar39last)

$sor401.Copy($tar401)                                        #Action for 40th row
$sor402.Copy($tar402)
$sor403.Copy($tar403)
$sor404.Copy($tar404)
$sor405.Copy($tar405)
$sor406.Copy($tar406)
$sor407.Copy($tar407)
$sor40last.Copy($tar40last)

$sor411.Copy($tar411)                                        #Action for 41th row
$sor412.Copy($tar412)
$sor413.Copy($tar413)
$sor414.Copy($tar414)
$sor415.Copy($tar415)
$sor416.Copy($tar416)
$sor417.Copy($tar417)
$sor41last.Copy($tar41last)

$sor421.Copy($tar421)                                        #Action for 42th row
$sor422.Copy($tar422)
$sor423.Copy($tar423)
$sor424.Copy($tar424)
$sor425.Copy($tar425)
$sor426.Copy($tar426)
$sor427.Copy($tar427)
$sor42last.Copy($tar42last)

$sor431.Copy($tar431)                                        #Action for 43th row
$sor432.Copy($tar432)
$sor433.Copy($tar433)
$sor434.Copy($tar434)
$sor435.Copy($tar435)
$sor436.Copy($tar436)
$sor437.Copy($tar437)
$sor43last.Copy($tar43last)

$sor441.Copy($tar441)                                        #Action for 44th row
$sor442.Copy($tar442)
$sor443.Copy($tar443)
$sor444.Copy($tar444)
$sor445.Copy($tar445)
$sor446.Copy($tar446)
$sor447.Copy($tar447)
$sor44last.Copy($tar44last)

$sor451.Copy($tar451)                                        #Action for 45th row
$sor452.Copy($tar452)
$sor453.Copy($tar453)
$sor454.Copy($tar454)
$sor455.Copy($tar455)
$sor456.Copy($tar456)
$sor457.Copy($tar457)
$sor45last.Copy($tar45last)

$sor461.Copy($tar461)                                        #Action for 46th row
$sor462.Copy($tar462)
$sor463.Copy($tar463)
$sor464.Copy($tar464)
$sor465.Copy($tar465)
$sor466.Copy($tar466)
$sor467.Copy($tar467)
$sor46last.Copy($tar46last)

$sor471.Copy($tar471)                                        #Action for 47th row
$sor472.Copy($tar472)
$sor473.Copy($tar473)
$sor474.Copy($tar474)
$sor475.Copy($tar475)
$sor476.Copy($tar476)
$sor477.Copy($tar477)
$sor47last.Copy($tar47last)

$sor481.Copy($tar481)                                        #Action for 48th row
$sor482.Copy($tar482)
$sor483.Copy($tar483)
$sor484.Copy($tar484)
$sor485.Copy($tar485)
$sor486.Copy($tar486)
$sor487.Copy($tar487)
$sor48last.Copy($tar48last)

$sor491.Copy($tar491)                                        #Action for 49th row
$sor492.Copy($tar492)
$sor493.Copy($tar493)
$sor494.Copy($tar494)
$sor495.Copy($tar495)
$sor496.Copy($tar496)
$sor497.Copy($tar497)
$sor49last.Copy($tar49last)

$sor501.Copy($tar501)                                        #Action for 50th row
$sor502.Copy($tar502)
$sor503.Copy($tar503)
$sor504.Copy($tar504)
$sor505.Copy($tar505)
$sor506.Copy($tar506)
$sor507.Copy($tar507)
$sor50last.Copy($tar50last)

$sor511.Copy($tar511)                                        #Action for 51th row
$sor512.Copy($tar512)
$sor513.Copy($tar513)
$sor514.Copy($tar514)
$sor515.Copy($tar515)
$sor516.Copy($tar516)
$sor517.Copy($tar517)
$sor51last.Copy($tar51last)

$sor521.Copy($tar521)                                        #Action for 52th row
$sor522.Copy($tar522)
$sor523.Copy($tar523)
$sor524.Copy($tar524)
$sor525.Copy($tar525)
$sor526.Copy($tar526)
$sor527.Copy($tar527)
$sor52last.Copy($tar52last)

$sor531.Copy($tar531)                                        #Action for 53th row
$sor532.Copy($tar532)
$sor533.Copy($tar533)
$sor534.Copy($tar534)
$sor535.Copy($tar535)
$sor536.Copy($tar536)
$sor537.Copy($tar537)
$sor53last.Copy($tar53last)

$sor541.Copy($tar541)                                        #Action for 54th row
$sor542.Copy($tar542)
$sor543.Copy($tar543)
$sor544.Copy($tar544)
$sor545.Copy($tar545)
$sor546.Copy($tar546)
$sor547.Copy($tar547)
$sor54last.Copy($tar54last)

$sor551.Copy($tar551)                                        #Action for 55th row
$sor552.Copy($tar552)
$sor553.Copy($tar553)
$sor554.Copy($tar554)
$sor555.Copy($tar555)
$sor556.Copy($tar556)
$sor557.Copy($tar557)
$sor55last.Copy($tar55last)

$sor561.Copy($tar561)                                        #Action for 56th row
$sor562.Copy($tar562)
$sor563.Copy($tar563)
$sor564.Copy($tar564)
$sor565.Copy($tar565)
$sor566.Copy($tar566)
$sor567.Copy($tar567)
$sor56last.Copy($tar56last)

$sor571.Copy($tar571)                                        #Action for 57th row
$sor572.Copy($tar572)
$sor573.Copy($tar573)
$sor574.Copy($tar574)
$sor575.Copy($tar575)
$sor576.Copy($tar576)
$sor577.Copy($tar577)
$sor57last.Copy($tar57last)

$sor581.Copy($tar581)                                        #Action for 58th row
$sor582.Copy($tar582)
$sor583.Copy($tar583)
$sor584.Copy($tar584)
$sor585.Copy($tar585)
$sor586.Copy($tar586)
$sor587.Copy($tar587)
$sor58last.Copy($tar58last)

$sor591.Copy($tar591)                                        #Action for 59th row
$sor592.Copy($tar592)
$sor593.Copy($tar593)
$sor594.Copy($tar594)
$sor595.Copy($tar595)
$sor596.Copy($tar596)
$sor597.Copy($tar597)
$sor59last.Copy($tar59last)

$sor601.Copy($tar601)                                        #Action for 60th row
$sor602.Copy($tar602)
$sor603.Copy($tar603)
$sor604.Copy($tar604)
$sor605.Copy($tar605)
$sor606.Copy($tar606)
$sor607.Copy($tar607)
$sor60last.Copy($tar60last)

$sor611.Copy($tar611)                                        #Action for 61th row
$sor612.Copy($tar612)
$sor613.Copy($tar613)
$sor614.Copy($tar614)
$sor615.Copy($tar615)
$sor616.Copy($tar616)
$sor617.Copy($tar617)
$sor61last.Copy($tar61last)

$sor621.Copy($tar621)                                        #Action for 62th row
$sor622.Copy($tar622)
$sor623.Copy($tar623)
$sor624.Copy($tar624)
$sor625.Copy($tar625)
$sor626.Copy($tar626)
$sor627.Copy($tar627)
$sor62last.Copy($tar62last)

$sor631.Copy($tar631)                                        #Action for 63th row
$sor632.Copy($tar632)
$sor633.Copy($tar633)
$sor634.Copy($tar634)
$sor635.Copy($tar635)
$sor636.Copy($tar636)
$sor637.Copy($tar637)
$sor63last.Copy($tar63last)

$sor641.Copy($tar641)                                        #Action for 64th row
$sor642.Copy($tar642)
$sor643.Copy($tar643)
$sor644.Copy($tar644)
$sor645.Copy($tar645)
$sor646.Copy($tar646)
$sor647.Copy($tar647)
$sor64last.Copy($tar64last)

$sor651.Copy($tar651)                                        #Action for 65th row
$sor652.Copy($tar652)
$sor653.Copy($tar653)
$sor654.Copy($tar654)
$sor655.Copy($tar655)
$sor656.Copy($tar656)
$sor657.Copy($tar657)
$sor65last.Copy($tar65last)

$sor661.Copy($tar661)                                        #Action for 66th row
$sor662.Copy($tar662)
$sor663.Copy($tar663)
$sor664.Copy($tar664)
$sor665.Copy($tar665)
$sor666.Copy($tar666)
$sor667.Copy($tar667)
$sor66last.Copy($tar66last)

$sor671.Copy($tar671)                                        #Action for 67th row
$sor672.Copy($tar672)
$sor673.Copy($tar673)
$sor674.Copy($tar674)
$sor675.Copy($tar675)
$sor676.Copy($tar676)
$sor677.Copy($tar677)
$sor67last.Copy($tar67last)

$sor681.Copy($tar681)                                        #Action for 68th row
$sor682.Copy($tar682)
$sor683.Copy($tar683)
$sor684.Copy($tar684)
$sor685.Copy($tar685)
$sor686.Copy($tar686)
$sor687.Copy($tar687)
$sor68last.Copy($tar68last)

$sor691.Copy($tar691)                                        #Action for 69th row
$sor692.Copy($tar692)
$sor693.Copy($tar693)
$sor694.Copy($tar694)
$sor695.Copy($tar695)
$sor696.Copy($tar696)
$sor697.Copy($tar697)
$sor69last.Copy($tar69last)

$sor701.Copy($tar701)                                        #Action for 70th row
$sor702.Copy($tar702)
$sor703.Copy($tar703)
$sor704.Copy($tar704)
$sor705.Copy($tar705)
$sor706.Copy($tar706)
$sor707.Copy($tar707)
$sor70last.Copy($tar70last)

$sor711.Copy($tar711)                                        #Action for 71th row
$sor712.Copy($tar712)
$sor713.Copy($tar713)
$sor714.Copy($tar714)
$sor715.Copy($tar715)
$sor716.Copy($tar716)
$sor717.Copy($tar717)
$sor71last.Copy($tar71last)

$sor721.Copy($tar721)                                        #Action for 72th row
$sor722.Copy($tar722)
$sor723.Copy($tar723)
$sor724.Copy($tar724)
$sor725.Copy($tar725)
$sor726.Copy($tar726)
$sor727.Copy($tar727)
$sor72last.Copy($tar72last)

$sor731.Copy($tar731)                                        #Action for 73th row
$sor732.Copy($tar732)
$sor733.Copy($tar733)
$sor734.Copy($tar734)
$sor735.Copy($tar735)
$sor736.Copy($tar736)
$sor737.Copy($tar737)
$sor73last.Copy($tar73last)

$sor741.Copy($tar741)                                        #Action for 74th row
$sor742.Copy($tar742)
$sor743.Copy($tar743)
$sor744.Copy($tar744)
$sor745.Copy($tar745)
$sor746.Copy($tar746)
$sor747.Copy($tar747)
$sor74last.Copy($tar74last)

$sor751.Copy($tar751)                                        #Action for 75th row
$sor752.Copy($tar752)
$sor753.Copy($tar753)
$sor754.Copy($tar754)
$sor755.Copy($tar755)
$sor756.Copy($tar756)
$sor757.Copy($tar757)
$sor75last.Copy($tar75last)

$sor761.Copy($tar761)                                        #Action for 76th row
$sor762.Copy($tar762)
$sor763.Copy($tar763)
$sor764.Copy($tar764)
$sor765.Copy($tar765)
$sor766.Copy($tar766)
$sor767.Copy($tar767)
$sor76last.Copy($tar76last)

$sor771.Copy($tar771)                                        #Action for 77th row
$sor772.Copy($tar772)
$sor773.Copy($tar773)
$sor774.Copy($tar774)
$sor775.Copy($tar775)
$sor776.Copy($tar776)
$sor777.Copy($tar777)
$sor77last.Copy($tar77last)

$sor781.Copy($tar781)                                        #Action for 78th row
$sor782.Copy($tar782)
$sor783.Copy($tar783)
$sor784.Copy($tar784)
$sor785.Copy($tar785)
$sor786.Copy($tar786)
$sor787.Copy($tar787)
$sor78last.Copy($tar78last)

$sor791.Copy($tar791)                                        #Action for 79th row
$sor792.Copy($tar792)
$sor793.Copy($tar793)
$sor794.Copy($tar794)
$sor795.Copy($tar795)
$sor796.Copy($tar796)
$sor797.Copy($tar797)
$sor79last.Copy($tar79last)

$sor801.Copy($tar801)                                        #Action for 80th row
$sor802.Copy($tar802)
$sor803.Copy($tar803)
$sor804.Copy($tar804)
$sor805.Copy($tar805)
$sor806.Copy($tar806)
$sor807.Copy($tar807)
$sor80last.Copy($tar80last)

$sor811.Copy($tar811)                                        #Action for 81th row
$sor812.Copy($tar812)
$sor813.Copy($tar813)
$sor814.Copy($tar814)
$sor815.Copy($tar815)
$sor816.Copy($tar816)
$sor817.Copy($tar817)
$sor81last.Copy($tar81last)

$sor821.Copy($tar821)                                        #Action for 82th row
$sor822.Copy($tar822)
$sor823.Copy($tar823)
$sor824.Copy($tar824)
$sor825.Copy($tar825)
$sor826.Copy($tar826)
$sor827.Copy($tar827)
$sor82last.Copy($tar82last)

$sor831.Copy($tar831)                                        #Action for 83th row
$sor832.Copy($tar832)
$sor833.Copy($tar833)
$sor834.Copy($tar834)
$sor835.Copy($tar835)
$sor836.Copy($tar836)
$sor837.Copy($tar837)
$sor83last.Copy($tar83last)

$sor841.Copy($tar841)                                        #Action for 84th row
$sor842.Copy($tar842)
$sor843.Copy($tar843)
$sor844.Copy($tar844)
$sor845.Copy($tar845)
$sor846.Copy($tar846)
$sor847.Copy($tar847)
$sor84last.Copy($tar84last)

$sor851.Copy($tar851)                                        #Action for 85th row
$sor852.Copy($tar852)
$sor853.Copy($tar853)
$sor854.Copy($tar854)
$sor855.Copy($tar855)
$sor856.Copy($tar856)
$sor857.Copy($tar857)
$sor85last.Copy($tar85last)

$sor861.Copy($tar861)                                        #Action for 86th row
$sor862.Copy($tar862)
$sor863.Copy($tar863)
$sor864.Copy($tar864)
$sor865.Copy($tar865)
$sor866.Copy($tar866)
$sor867.Copy($tar867)
$sor86last.Copy($tar86last)

$sor871.Copy($tar871)                                        #Action for 87th row
$sor872.Copy($tar872)
$sor873.Copy($tar873)
$sor874.Copy($tar874)
$sor875.Copy($tar875)
$sor876.Copy($tar876)
$sor877.Copy($tar877)
$sor87last.Copy($tar87last)

$sor881.Copy($tar881)                                        #Action for 88th row
$sor882.Copy($tar882)
$sor883.Copy($tar883)
$sor884.Copy($tar884)
$sor885.Copy($tar885)
$sor886.Copy($tar886)
$sor887.Copy($tar887)
$sor88last.Copy($tar88last)

$sor891.Copy($tar891)                                        #Action for 89th row
$sor892.Copy($tar892)
$sor893.Copy($tar893)
$sor894.Copy($tar894)
$sor895.Copy($tar895)
$sor896.Copy($tar896)
$sor897.Copy($tar897)
$sor89last.Copy($tar89last)

$sor901.Copy($tar901)                                        #Action for 90th row
$sor902.Copy($tar902)
$sor903.Copy($tar903)
$sor904.Copy($tar904)
$sor905.Copy($tar905)
$sor906.Copy($tar906)
$sor907.Copy($tar907)
$sor90last.Copy($tar90last)

$sor911.Copy($tar911)                                        #Action for 91th row
$sor912.Copy($tar912)
$sor913.Copy($tar913)
$sor914.Copy($tar914)
$sor915.Copy($tar915)
$sor916.Copy($tar916)
$sor917.Copy($tar917)
$sor91last.Copy($tar91last)

$sor921.Copy($tar921)                                        #Action for 92th row
$sor922.Copy($tar922)
$sor923.Copy($tar923)
$sor924.Copy($tar924)
$sor925.Copy($tar925)
$sor926.Copy($tar926)
$sor927.Copy($tar927)
$sor92last.Copy($tar92last)

$sor931.Copy($tar931)                                        #Action for 93th row
$sor932.Copy($tar932)
$sor933.Copy($tar933)
$sor934.Copy($tar934)
$sor935.Copy($tar935)
$sor936.Copy($tar936)
$sor937.Copy($tar937)
$sor93last.Copy($tar93last)

$sor941.Copy($tar941)                                        #Action for 94th row
$sor942.Copy($tar942)
$sor943.Copy($tar943)
$sor944.Copy($tar944)
$sor945.Copy($tar945)
$sor946.Copy($tar946)
$sor947.Copy($tar947)
$sor94last.Copy($tar94last)

$sor951.Copy($tar951)                                        #Action for 95th row
$sor952.Copy($tar952)
$sor953.Copy($tar953)
$sor954.Copy($tar954)
$sor955.Copy($tar955)
$sor956.Copy($tar956)
$sor957.Copy($tar957)
$sor95last.Copy($tar95last)

$sor961.Copy($tar961)                                        #Action for 96th row
$sor962.Copy($tar962)
$sor963.Copy($tar963)
$sor964.Copy($tar964)
$sor965.Copy($tar965)
$sor966.Copy($tar966)
$sor967.Copy($tar967)
$sor96last.Copy($tar96last)

$sor971.Copy($tar971)                                        #Action for 97th row
$sor972.Copy($tar972)
$sor973.Copy($tar973)
$sor974.Copy($tar974)
$sor975.Copy($tar975)
$sor976.Copy($tar976)
$sor977.Copy($tar977)
$sor97last.Copy($tar97last)

$sor981.Copy($tar981)                                        #Action for 98th row
$sor982.Copy($tar982)
$sor983.Copy($tar983)
$sor984.Copy($tar984)
$sor985.Copy($tar985)
$sor986.Copy($tar986)
$sor987.Copy($tar987)
$sor98last.Copy($tar98last)

$sor991.Copy($tar991)                                        #Action for 99th row
$sor992.Copy($tar992)
$sor993.Copy($tar993)
$sor994.Copy($tar994)
$sor995.Copy($tar995)
$sor996.Copy($tar996)
$sor997.Copy($tar997)
$sor99last.Copy($tar99last)

$sor1001.Copy($tar1001)                                        #Action for 100th row
$sor1002.Copy($tar1002)
$sor1003.Copy($tar1003)
$sor1004.Copy($tar1004)
$sor1005.Copy($tar1005)
$sor1006.Copy($tar1006)
$sor1007.Copy($tar1007)
$sor100last.Copy($tar100last)
#------------------------------------------------------------------------------------------
$wb2.close($false) # close source workbook w/o saving 
$wb.close($true) # close and save destination workbook 
$xl.quit()
 [System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)