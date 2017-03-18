
$pageURL = "http://www.pornhub.com/video/search?search=bbc+interracial&page=";
$pageCount = 2;
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
Write-Host $aLinks.Count;
$aLinks | ForEach-Object {
    $Link = $_;
    if($Link.HREF -like '*view_video*'){
        if($Links.Contains(($Link.href)) -eq $false){
            $Links += , ($Link.href);
        }
            #$xe.Navigate($Link);
            #while ($xe.Busy) {
            #    [System.Threading.Thread]::Sleep(10)
            #}
            #[System.Threading.Thread]::Sleep(3000);
            #$xe.Document.Script.execScript("document.querySelector('.js-voteDown').click();","JavaScript");
            #$xe.Stop();
                
    }
}
$Links | ForEach-Object{
    Write-Host $_;
    $xe.Navigate($_);
    while ($xe.Busy) {
        [System.Threading.Thread]::Sleep(10)
    }
    [System.Threading.Thread]::Sleep(4000);
    $xe.Document.Script.execScript("document.querySelector('.js-voteDown').click();","JavaScript");
}
$xe.Navigate("http://Google.com");
#[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xe)
