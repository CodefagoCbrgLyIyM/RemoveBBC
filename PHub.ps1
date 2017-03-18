
#Page url with search tearms
$pageURL = "http://www.pornhub.com/video/search?search=bbc+interracial&page=";
#Amount of pages to scan
$pageCount = 5;


$xe = New-Object -com "InternetExplorer.Application";
$xe.visible = $true;
$xe.silent = $true;
$aLinks = @();
$Links = @();
for($i = 1; $i -le $pageCount; $i++){
    $URL = $pageURL + $i.ToString();
    $xe.Navigate($URL)
    while ($xe.Busy) {
        [System.Threading.Thread]::Sleep(10)
    } 
    
    $aLinks += $xe.Document.IHTMLDocument3_getElementsByTagName('a');
    $il = 0
    $xe.Stop();
    
}
Write-Host $aLinks.Count "Links found in " $pageCount " pages.";
$aLinks | ForEach-Object {
    $Link = $_;
    if($Link.HREF -like '*view_video*'){
        if(($Links.Contains(($Link.href)) -eq $false) -and ($Link.title -eq $Link.innerText)){
            $Links += , ($Link.href);
        }                
    }
}
Write-Host $Links.Count " Videos to downvote";
$Links | ForEach-Object{
    
    $xe.Navigate($_);
    while ($xe.Busy) {
        [System.Threading.Thread]::Sleep(10)
    }
    [System.Threading.Thread]::Sleep(4000);
    $xe.Document.Script.execScript("document.querySelector('.js-voteDown').click();","JavaScript");
    Write-Host "Downvoted " $_;
}
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xe);
