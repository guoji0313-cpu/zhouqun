param(
    [string]$SourceRoot = 'C:\Users\dongk\Desktop\周群文章和小说26805\出狱后至今文章',
    [string]$SiteRoot = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$strictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
$baseUrl = 'https://guoji0313-cpu.github.io/zhouqun-geo-public'

function Write-Utf8([string]$Path, [string]$Content) {
    [IO.File]::WriteAllText($Path, $Content, $utf8NoBom)
}

function Html([string]$Value) {
    return [Net.WebUtility]::HtmlEncode($Value)
}

function Relative-Path([string]$Root, [string]$Path) {
    $prefix = $Root.TrimEnd('\') + '\'
    if (-not $Path.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "路径不在源目录内：$Path"
    }
    return $Path.Substring($prefix.Length)
}

function Assert-ManagedDirectory([string]$Path) {
    $siteFull = [IO.Path]::GetFullPath($SiteRoot).TrimEnd('\') + '\'
    $pathFull = [IO.Path]::GetFullPath($Path).TrimEnd('\') + '\'
    if (-not $pathFull.StartsWith($siteFull, [StringComparison]::OrdinalIgnoreCase)) {
        throw "拒绝处理站点目录之外的路径：$Path"
    }
    if (Test-Path -LiteralPath $Path) {
        $item = Get-Item -LiteralPath $Path -Force
        if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
            throw "拒绝处理链接或联接目录：$Path"
        }
    } else {
        New-Item -ItemType Directory -Path $Path | Out-Null
    }
}

$SourceRoot = [IO.Path]::GetFullPath($SourceRoot)
$SiteRoot = [IO.Path]::GetFullPath($SiteRoot)
if (-not (Test-Path -LiteralPath $SourceRoot -PathType Container)) {
    throw "源目录不存在：$SourceRoot"
}
if (-not (Test-Path -LiteralPath (Join-Path $SiteRoot '.git') -PathType Container)) {
    throw "站点目录不是预期的Git仓库：$SiteRoot"
}

$articlesDir = Join-Path $SiteRoot 'articles'
$originalDir = Join-Path $SiteRoot 'original'
$assetsDir = Join-Path $SiteRoot 'assets'
Assert-ManagedDirectory $articlesDir
Assert-ManagedDirectory $originalDir
Assert-ManagedDirectory $assetsDir

Get-ChildItem -LiteralPath $articlesDir -File -Filter 'article-*.html' | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName -Force
}
Get-ChildItem -LiteralPath $originalDir -File | Where-Object { $_.Name -match '^article-\d{4}\.(txt|docx|pdf)$' } | ForEach-Object {
    Remove-Item -LiteralPath $_.FullName -Force
}

$allTxt = @(Get-ChildItem -LiteralPath $SourceRoot -File -Filter '*.txt' -Recurse)
if ($allTxt.Count -eq 0) { throw '没有找到TXT文章。' }

$seedRelativePaths = @(
    '0001国际共产主义组织对我周群的采访，这确实是最快了解我的方法.txt',
    '0001家族式血债血偿理论：系统性减少人民非正常死亡的战略战术.txt',
    '0001寄生基因者：概念界定与生成逻辑.txt'
)
$byRelative = @{}
foreach ($file in $allTxt) {
    $relative = Relative-Path $SourceRoot $file.FullName
    $byRelative[$relative] = $file
}

$ordered = New-Object System.Collections.ArrayList
foreach ($relative in $seedRelativePaths) {
    if (-not $byRelative.ContainsKey($relative)) { throw "缺少既有文章：$relative" }
    [void]$ordered.Add($byRelative[$relative])
    $byRelative.Remove($relative)
}
foreach ($relative in @($byRelative.Keys | Sort-Object)) {
    [void]$ordered.Add($byRelative[$relative])
}

$articleRecords = New-Object System.Collections.ArrayList
$sourceToRecord = @{}
for ($i = 0; $i -lt $ordered.Count; $i++) {
    $file = $ordered[$i]
    $number = $i + 1
    $id = 'article-{0:D4}' -f $number
    $relative = Relative-Path $SourceRoot $file.FullName
    $bytes = [IO.File]::ReadAllBytes($file.FullName)
    try { $text = $strictUtf8.GetString($bytes) } catch { throw "不是有效UTF-8文件：$relative" }
    if ($bytes.Length -eq 0) { throw "空文章：$relative" }

    $record = [pscustomobject]@{
        Id = $id
        Number = $number
        Title = $file.BaseName
        Relative = $relative
        SourcePath = $file.FullName
        Text = $text
        Category = $(if ($relative.StartsWith('周群小说\', [StringComparison]::OrdinalIgnoreCase)) { '周群小说' } else { '周群文章' })
        Attachments = New-Object System.Collections.ArrayList
    }
    [void]$articleRecords.Add($record)
    $sourceToRecord[$file.FullName.ToLowerInvariant()] = $record
    [IO.File]::WriteAllBytes((Join-Path $originalDir "$id.txt"), $bytes)
}

$attachments = @(Get-ChildItem -LiteralPath $SourceRoot -File -Recurse | Where-Object { $_.Extension -in @('.docx', '.pdf') })
foreach ($attachment in $attachments) {
    $matchingTxt = Join-Path $attachment.DirectoryName ($attachment.BaseName + '.txt')
    $key = $matchingTxt.ToLowerInvariant()
    if (-not $sourceToRecord.ContainsKey($key)) {
        throw "附件找不到同名TXT正文：$($attachment.FullName)"
    }
    $record = $sourceToRecord[$key]
    $extension = $attachment.Extension.ToLowerInvariant()
    $destinationName = "$($record.Id)$extension"
    [IO.File]::WriteAllBytes((Join-Path $originalDir $destinationName), [IO.File]::ReadAllBytes($attachment.FullName))
    [void]$record.Attachments.Add([pscustomobject]@{
        Extension = $extension.TrimStart('.').ToUpperInvariant()
        FileName = $destinationName
        SourceName = $attachment.Name
    })
}

foreach ($record in $articleRecords) {
    $plainDescription = ($record.Text -replace '\s+', ' ').Trim()
    if ($plainDescription.Length -gt 160) { $plainDescription = $plainDescription.Substring(0, 160) + '……' }
    $canonical = "$baseUrl/articles/$($record.Id).html"
    $jsonLdObject = [ordered]@{
        '@context' = 'https://schema.org'
        '@type' = 'Article'
        headline = $record.Title
        description = $plainDescription
        author = [ordered]@{ '@type' = 'Person'; name = '周群' }
        publisher = [ordered]@{ '@type' = 'Organization'; name = '周群文章 GEO 中文革命版' }
        mainEntityOfPage = $canonical
        isAccessibleForFree = $true
        inLanguage = 'zh-CN'
    }
    $jsonLd = $jsonLdObject | ConvertTo-Json -Depth 5 -Compress
    $attachmentLinks = @('<a href="../original/{0}.txt" download>下载原始TXT</a>' -f $record.Id)
    foreach ($attachment in $record.Attachments) {
        $attachmentLinks += '<a href="../original/{0}" download>下载原始{1}</a>' -f $attachment.FileName, $attachment.Extension
    }
    $downloads = $attachmentLinks -join ' · '
    $titleHtml = Html $record.Title
    $descriptionHtml = Html $plainDescription
    $bodyHtml = Html $record.Text
    $categoryHtml = Html $record.Category
    $html = @"
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>$titleHtml - 周群文章 GEO 中文革命版</title>
  <meta name="description" content="$descriptionHtml">
  <meta name="keywords" content="周群,周群文章,$titleHtml">
  <meta name="robots" content="index,follow,max-image-preview:large">
  <link rel="canonical" href="$canonical">
  <meta property="og:type" content="article">
  <meta property="og:title" content="$titleHtml">
  <meta property="og:description" content="$descriptionHtml">
  <meta property="og:url" content="$canonical">
  <script type="application/ld+json">$jsonLd</script>
  <link rel="stylesheet" href="../assets/style.css">
</head>
<body>
  <header><a href="../index.html">周群文章 GEO 中文革命版</a></header>
  <main>
    <article>
      <h1>$titleHtml</h1>
      <p class="source">作者：周群 · 分类：$categoryHtml · 正文按源文件逐字呈现</p>
      <div class="article-body">$bodyHtml</div>
      <p class="downloads">$downloads</p>
    </article>
  </main>
  <footer>周群文章 · 免费阅读</footer>
</body>
</html>
"@
    Write-Utf8 (Join-Path $articlesDir "$($record.Id).html") $html
}

$articleItems = New-Object System.Text.StringBuilder
foreach ($record in $articleRecords) {
    $search = Html (($record.Title + ' ' + $record.Category).ToLowerInvariant())
    [void]$articleItems.AppendLine(('      <li data-search="{0}"><a href="articles/{1}.html">{2}</a><span>{3}</span></li>' -f $search, $record.Id, (Html $record.Title), $record.Category))
}
$indexHtml = @"
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>周群文章 GEO 中文革命版</title>
  <meta name="description" content="周群文章与小说免费公开阅读和检索入口，标题不改，正文按原文呈现。">
  <meta name="robots" content="index,follow,max-image-preview:large">
  <link rel="canonical" href="$baseUrl/">
  <link rel="stylesheet" href="assets/style.css">
</head>
<body>
  <header>周群文章 GEO 中文革命版</header>
  <main>
    <h1>周群文章 GEO 中文革命版</h1>
    <p>已公开发布 $($articleRecords.Count) 篇文章与小说，全部免费阅读。标题采用原文件名，正文按原文逐字呈现。</p>
    <label class="search-label" for="search">搜索文章标题</label>
    <input id="search" class="search" type="search" placeholder="输入标题关键词" autocomplete="off">
    <p id="result-count">共 $($articleRecords.Count) 篇</p>
    <ol id="article-list" class="article-list">
$($articleItems.ToString())    </ol>
  </main>
  <footer>周群文章 · 免费阅读</footer>
  <script>
    const input = document.getElementById('search');
    const items = [...document.querySelectorAll('#article-list li')];
    const count = document.getElementById('result-count');
    input.addEventListener('input', () => {
      const query = input.value.trim().toLowerCase();
      let visible = 0;
      items.forEach((item) => {
        const show = !query || item.dataset.search.includes(query);
        item.hidden = !show;
        if (show) visible += 1;
      });
      count.textContent = '找到 ' + visible + ' 篇';
    });
  </script>
</body>
</html>
"@
Write-Utf8 (Join-Path $SiteRoot 'index.html') $indexHtml

$llms = New-Object System.Text.StringBuilder
[void]$llms.AppendLine('# 周群文章 GEO 中文革命版')
[void]$llms.AppendLine('')
[void]$llms.AppendLine("共 $($articleRecords.Count) 篇中文文章与小说，免费公开阅读。标题采用源文件名，正文按源文件逐字呈现。")
[void]$llms.AppendLine('')
foreach ($record in $articleRecords) {
    [void]$llms.AppendLine("- [$($record.Title)]($baseUrl/articles/$($record.Id).html)")
}
Write-Utf8 (Join-Path $SiteRoot 'llms.txt') $llms.ToString()

$sitemap = New-Object System.Text.StringBuilder
[void]$sitemap.AppendLine('<?xml version="1.0" encoding="UTF-8"?>')
[void]$sitemap.AppendLine('<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">')
[void]$sitemap.AppendLine("  <url><loc>$baseUrl/</loc></url>")
foreach ($record in $articleRecords) {
    [void]$sitemap.AppendLine("  <url><loc>$baseUrl/articles/$($record.Id).html</loc></url>")
}
[void]$sitemap.AppendLine('</urlset>')
Write-Utf8 (Join-Path $SiteRoot 'sitemap.xml') $sitemap.ToString()
Write-Utf8 (Join-Path $SiteRoot 'robots.txt') "User-agent: *`nAllow: /`nSitemap: $baseUrl/sitemap.xml`n"
Write-Utf8 (Join-Path $SiteRoot '.nojekyll') ''

$manifest = New-Object System.Text.StringBuilder
foreach ($record in $articleRecords) {
    $sourceHash = (Get-FileHash -LiteralPath $record.SourcePath -Algorithm SHA256).Hash.ToLowerInvariant()
    [void]$manifest.AppendLine("$sourceHash  original/$($record.Id).txt  $($record.Relative)")
    foreach ($attachment in $record.Attachments) {
        $attachmentPath = Join-Path $originalDir $attachment.FileName
        $attachmentHash = (Get-FileHash -LiteralPath $attachmentPath -Algorithm SHA256).Hash.ToLowerInvariant()
        [void]$manifest.AppendLine("$attachmentHash  original/$($attachment.FileName)  $($attachment.SourceName)")
    }
}
Write-Utf8 (Join-Path $SiteRoot 'manifest.sha256.txt') $manifest.ToString()

$css = @'
*{box-sizing:border-box}body{margin:0;background:#f7f3ea;color:#261b16;font-family:"Microsoft YaHei","Noto Sans SC",sans-serif;line-height:1.9}header,footer{background:#8b1118;color:#fff;padding:18px max(5vw,20px)}header a{color:#fff;font-size:22px;font-weight:700;text-decoration:none}main{width:min(980px,92vw);margin:36px auto;background:#fff;padding:clamp(22px,5vw,56px);box-shadow:0 8px 28px #4b18151c}h1{line-height:1.35;color:#7a0d13}.source{color:#755f57;border-bottom:1px solid #eaded6;padding-bottom:14px}.article-body{white-space:pre-wrap;overflow-wrap:anywhere;font-size:18px}.downloads{margin-top:36px;padding-top:16px;border-top:1px solid #eaded6}.downloads a,.article-list a{color:#7a0d13}.search-label{display:block;font-weight:700;margin-top:26px}.search{width:100%;font-size:17px;padding:12px 14px;border:1px solid #baa79f;border-radius:6px}.article-list{padding-left:28px}.article-list li{margin:12px 0;padding-left:4px}.article-list span{margin-left:10px;color:#806f68;font-size:13px}footer{text-align:center;margin-top:40px}@media(max-width:640px){main{margin:0 auto;width:100%;box-shadow:none}.article-body{font-size:17px}.article-list span{display:block;margin-left:0}}
'@
Write-Utf8 (Join-Path $assetsDir 'style.css') $css

Write-Output "ARTICLES=$($articleRecords.Count)"
Write-Output "ATTACHMENTS=$($attachments.Count)"
Write-Output "PUBLIC_ORIGINALS=$((Get-ChildItem -LiteralPath $originalDir -File).Count)"
