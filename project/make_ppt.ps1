$ErrorActionPreference = 'Stop'

$out = Join-Path $PSScriptRoot 'china-sns-landscape.pptx'
$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add()
$pres.PageSetup.SlideWidth = 13.333 * 72
$pres.PageSetup.SlideHeight = 7.5 * 72

$blank = 12
$black = 0x0A0A0A
$red = 0x4224FF
$pink = 0xE8E3FF
$lightPink = 0xD8D6FF
$gray = 0x6A6A6A
$white = 0xFFFFFF

function Add-Text($slide, $text, $x, $y, $w, $h, $size, $color, $font = 'Malgun Gothic', $bold = $false, $align = 1) {
  $box = $slide.Shapes.AddTextbox(1, $x, $y, $w, $h)
  $box.TextFrame.TextRange.Text = $text
  $box.TextFrame.TextRange.Font.Name = [string]$font
  $box.TextFrame.TextRange.Font.Size = $size
  $box.TextFrame.TextRange.Font.Bold = if ($bold) { -1 } else { 0 }
  $box.TextFrame.TextRange.Font.Color.RGB = $color
  $box.TextFrame.WordWrap = -1
  $box.TextFrame.MarginLeft = 0
  $box.TextFrame.MarginRight = 0
  $box.TextFrame.MarginTop = 0
  $box.TextFrame.MarginBottom = 0
  return $box
}

function Add-Rect($slide, $x, $y, $w, $h, $fill, $radius = $false, $line = $fill) {
  $kind = if ($radius) { 5 } else { 1 }
  $shape = $slide.Shapes.AddShape($kind, $x, $y, $w, $h)
  $shape.Fill.ForeColor.RGB = $fill
  $shape.Line.ForeColor.RGB = $line
  $shape.Line.Weight = 0.75
  return $shape
}

function Add-Circle($slide, $x, $y, $d, $fill, $line = $fill) {
  $shape = $slide.Shapes.AddShape(9, $x, $y, $d, $d)
  $shape.Fill.ForeColor.RGB = $fill
  $shape.Line.ForeColor.RGB = $line
  return $shape
}

function Add-IconBadge($slide, $label, $x, $y, $fill = $white, $color = $red) {
  $circle = Add-Circle $slide $x $y 34 $fill $fill
  Add-Text $slide $label $x $($y + 7) 34 18 13 $color 'Malgun Gothic' $true 2 | Out-Null
}

function Add-Footer($slide, $page, $source) {
  Add-Text $slide "PAGE $page / 05" 64 705 170 18 10 $gray 'Consolas' $false 1 | Out-Null
  Add-Text $slide $source 870 705 350 18 10 $gray 'Consolas' $false 3 | Out-Null
}

# Cover
$s = $pres.Slides.Add($pres.Slides.Count + 1, $blank)
$s.Background.Fill.ForeColor.RGB = 0xFFFFFF
Add-Text $s 'MIRI COMPANY  ·  CHINA SNS REPORT' 70 42 430 20 11 0x5A568C 'Malgun Gothic' $true 1 | Out-Null
Add-Text $s 'CHINA SOCIAL MEDIA' 430 93 470 22 12 $gray 'Consolas' $false 2 | Out-Null
Add-Text $s '중국 SNS\n종류 · 특징 · 문화' 245 125 840 116 42 0x302C72 'Malgun Gothic' $true 2 | Out-Null
Add-Text $s '중국인의 일상과 소비를 움직이는 디지털 플랫폼을 한눈에 읽다' 300 255 720 30 16 $gray 'Malgun Gothic' $false 2 | Out-Null

# Phone mockup and surrounding platform signals, inspired by the reference cover.
Add-Circle $s 235 365 105 0xE9F4FF 0xE9F4FF | Out-Null
Add-Text $s '微信' 252 398 72 24 18 0x1677FF 'Malgun Gothic' $true 2 | Out-Null
Add-Circle $s 1045 360 108 0xFFE3F1 0xFFE3F1 | Out-Null
Add-Text $s 'RED' 1062 397 74 24 18 0xE54873 'Consolas' $true 2 | Out-Null
Add-Circle $s 210 555 88 0xFFF0D7 0xFFF0D7 | Out-Null
Add-Text $s '微博' 222 584 64 22 16 0xD12D43 'Malgun Gothic' $true 2 | Out-Null
Add-Circle $s 1100 550 88 0xE9E4FF 0xE9E4FF | Out-Null
Add-Text $s '抖音' 1118 580 52 22 16 0x302C72 'Malgun Gothic' $true 2 | Out-Null
$phone = Add-Rect $s 515 340 390 380 0x161A2A $true 0x161A2A
Add-Rect $s 532 365 356 333 0xF8F9FC $true 0xF8F9FC | Out-Null
Add-Rect $s 660 350 100 10 0x161A2A $true 0x161A2A | Out-Null
Add-Text $s 'SOCIAL' 575 405 270 34 25 0x302C72 'Consolas' $true 2 | Out-Null
Add-Text $s '중국 SNS 플랫폼 지도' 575 447 270 28 15 $black 'Malgun Gothic' $true 2 | Out-Null
Add-Rect $s 575 510 130 56 0xEEE9FF $true 0xEEE9FF | Out-Null
Add-Text $s '플랫폼' 575 527 130 18 14 0x302C72 'Malgun Gothic' $true 2 | Out-Null
Add-Rect $s 725 510 130 56 0xFFE8EC $true 0xFFE8EC | Out-Null
Add-Text $s '문화' 725 527 130 18 14 $red 'Malgun Gothic' $true 2 | Out-Null
Add-Text $s '2026  ·  CAPSTONE DESIGN' 470 775 480 18 11 $gray 'Consolas' $false 2 | Out-Null

# Platforms
$s = $pres.Slides.Add($pres.Slides.Count + 1, $blank)
$s.Background.Fill.ForeColor.RGB = 0xF9FBFB
Add-Text $s '01 — PLATFORMS' 64 52 400 24 13 $red 'Consolas' $false 1 | Out-Null
Add-Text $s '6대 중국 SNS, 한눈에 보기' 64 92 850 52 31 $black 'Malgun Gothic' $true 1 | Out-Null
$platforms = @(
  @('W','위챗','WECHAT','슈퍼앱 · 메신저+결제+미니프로그램','13억+','MAU'),
  @('WB','웨이보','WEIBO','공론장 · 실시간 트렌드/연예 이슈','6억+','MAU'),
  @('DY','더우인','DOUYIN','숏폼 · 알고리즘 기반 라이브 커머스','7.5억+','DAU'),
  @('XHS','샤오홍슈','RED','라이프스타일 · Z세대 검색엔진','3억+','MAU'),
  @('B','빌리빌리','BILIBILI','영상 커뮤니티 · 서브컬처와 탄막','3.4억+','MAU'),
  @('ZH','즈후','ZHIHU','지식 Q&A · 전문가형 콘텐츠','1억+','MAU')
)
for ($i = 0; $i -lt 6; $i++) {
  $col = $i % 3; $row = [math]::Floor($i / 3)
  $x = 64 + ($col * 390); $y = 175 + ($row * 230)
  Add-Rect $s $x $y 360 198 $white $true 0xE6E6E6 | Out-Null
  Add-IconBadge $s $platforms[$i][0] ($x + 24) ($y + 24) $pink $red
  Add-Text $s $platforms[$i][1] ($x + 72) ($y + 28) 145 24 18 $black 'Malgun Gothic' $true 1 | Out-Null
  Add-Text $s "($($platforms[$i][2]))" ($x + 218) ($y + 31) 110 18 10 0x8A8A8A 'Malgun Gothic' $true 1 | Out-Null
  Add-Text $s $platforms[$i][3] ($x + 24) ($y + 91) 310 28 12 $black 'Malgun Gothic' $true 1 | Out-Null
  Add-Text $s $platforms[$i][4] ($x + 24) ($y + 140) 175 35 29 $black 'Malgun Gothic' $true 1 | Out-Null
  Add-Text $s $platforms[$i][5] ($x + 205) ($y + 151) 80 20 11 0x8A8A8A 'Consolas' $false 1 | Out-Null
}
Add-Footer $s '02' 'SOURCE: QUESTMOBILE, 2025'

# Culture
$s = $pres.Slides.Add($pres.Slides.Count + 1, $blank)
$s.Background.Fill.ForeColor.RGB = 0xF9FBFB
Add-Circle $s 1180 675 150 $lightPink $lightPink | Out-Null
Add-Text $s '03' 68 58 42 42 17 $white 'Malgun Gothic' $true 2 | Out-Null
$s.Shapes.Item($s.Shapes.Count).Fill.ForeColor.RGB = $red
Add-Text $s '중국 SNS 문화를 읽는 4가지 키워드' 128 55 750 37 24 $black 'Malgun Gothic' $true 1 | Out-Null
Add-Text $s 'READING CHINESE SNS CULTURE' 128 100 520 20 11 $gray 'Consolas' $false 1 | Out-Null
$keys = @(
  @('⌂','KEYWORD 01','슈퍼앱`n생태계','위챗 하나로 메시지, 결제, 쇼핑까지 끝나는 생활 인프라.'),
  @('★','KEYWORD 02','왕훙(网红)`n경제','인플루언서가 곧 브랜드, 팔로워가 곧 매출로 이어지는 구조.'),
  @('▶','KEYWORD 03','라이브`n커머스','실시간 방송으로 소개하고 즉시 결제까지 끝내는 쇼핑.'),
  @('▣','KEYWORD 04','폐쇄형`n플랫폼','만리방화벽 안에서 독자적인 생태계가 빠르게 순환한다.')
)
for ($i = 0; $i -lt 4; $i++) {
  $x = 64 + ($i * 304); $y = 190 + (($i % 2) * 25)
  Add-Rect $s $x $y 278 380 $red $true $red | Out-Null
  Add-Text $s $keys[$i][0] ($x + 24) ($y + 25) 42 35 22 $white 'Malgun Gothic' $true 1 | Out-Null
  Add-Text $s $keys[$i][1] ($x + 24) ($y + 90) 190 18 10 $lightPink 'Consolas' $false 1 | Out-Null
  Add-Text $s $keys[$i][2] ($x + 24) ($y + 125) 220 75 24 $white 'Malgun Gothic' $true 1 | Out-Null
  Add-Text $s $keys[$i][3] ($x + 24) ($y + 280) 225 58 11 $pink 'Malgun Gothic' $true 1 | Out-Null
}
Add-Footer $s '03' 'KEYWORDS · WECHAT · DOUYIN · LIVE COMMERCE'

# Daily flow and implications
$s = $pres.Slides.Add($pres.Slides.Count + 1, $blank)
$s.Background.Fill.ForeColor.RGB = 0xF9FBFB
Add-Text $s '04 — DAILY FLOW' 64 52 400 24 13 $red 'Consolas' $false 1 | Out-Null
Add-Text $s '하루의 동선이 곧 플랫폼 생태계가 된다' 64 92 950 52 31 $black 'Malgun Gothic' $true 1 | Out-Null
Add-Text $s '중국 SNS는 콘텐츠 앱을 넘어 검색·결제·구매가 이어지는 생활 인프라다.' 66 151 1050 28 16 $gray 'Malgun Gothic' $false 1 | Out-Null
$flow = @(
  @('08:00','발견','샤오홍슈','후기·추천 검색'),
  @('12:30','소통','위챗','메시지·결제'),
  @('18:00','참여','빌리빌리','영상·탄막 커뮤니티'),
  @('21:00','전환','더우인','라이브·즉시 구매')
)
for ($i = 0; $i -lt 4; $i++) {
  $x = 64 + ($i * 300)
  Add-Rect $s $x 245 250 240 $white $true 0xE6E6E6 | Out-Null
  Add-Text $s $flow[$i][0] ($x + 24) 270 180 24 13 $red 'Consolas' $true 1 | Out-Null
  Add-Text $s $flow[$i][1] ($x + 24) 316 190 36 25 $black 'Malgun Gothic' $true 1 | Out-Null
  Add-Text $s $flow[$i][2] ($x + 24) 365 190 30 20 18 0x302C72 'Malgun Gothic' $true 1 | Out-Null
  Add-Text $s $flow[$i][3] ($x + 24) 425 195 22 12 $gray 'Malgun Gothic' $false 1 | Out-Null
  if ($i -lt 3) { Add-Text $s '→' ($x + 258) 350 30 30 20 22 $red 'Malgun Gothic' $true 2 | Out-Null }
}
Add-Rect $s 64 555 1120 82 0x302C72 $true 0x302C72 | Out-Null
Add-Text $s '핵심 시사점' 92 578 150 24 15 $white 'Malgun Gothic' $true 1 | Out-Null
Add-Text $s '중국 SNS를 이해하려면 “어떤 앱인가”보다 “어떤 생활 단계와 연결되는가”를 봐야 한다.' 265 574 850 32 17 $white 'Malgun Gothic' $true 1 | Out-Null
Add-Footer $s '04' 'SOURCE: PLATFORM REPORTS, 2025'

# Closing
$s = $pres.Slides.Add($pres.Slides.Count + 1, $blank)
$s.Background.Fill.ForeColor.RGB = 0xF9FBFB
$oval = $s.Shapes.AddShape(9, 920, 110, 510, 510); $oval.Fill.Visible = 0; $oval.Line.ForeColor.RGB = $lightPink; $oval.Line.Weight = 8
$oval = $s.Shapes.AddShape(9, 1040, 330, 360, 360); $oval.Fill.Visible = 0; $oval.Line.ForeColor.RGB = $red; $oval.Line.Weight = 8
$oval = $s.Shapes.AddShape(9, 1040, 565, 260, 260); $oval.Fill.Visible = 0; $oval.Line.ForeColor.RGB = 0x1F0F6B; $oval.Line.Weight = 8
Add-Text $s 'CHINA SNS LANDSCAPE REPORT' 64 105 600 24 13 $red 'Consolas' $true 1 | Out-Null
Add-Text $s "시청해주셔서`n감사합니다!" 64 210 750 125 42 $black 'Malgun Gothic' $true 1 | Out-Null
Add-Text $s '중국 SNS 지형도' 105 625 300 25 14 $black 'Malgun Gothic' $true 1 | Out-Null
Add-Text $s "PAGE 05 / 05 - SOURCE QUESTMOBILE 2025" 105 655 500 20 10 $gray "Consolas" $false 1 | Out-Null
Add-Rect $s 64 630 26 26 $red $false $red | Out-Null

$pres.SaveAs($out)
$pres.Close()
$pp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
Write-Output $out