# PowerShell 7 + desktop Microsoft Word on Windows. No input document is opened in Word.
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string[]]$Inputs,
    [Parameter(Mandatory)][ValidateNotNullOrEmpty()][string]$OutputDirectory,
    [string[]]$CaptionLabels = @('Figure', 'Table', '图', '表'),
    [switch]$PreflightOnly
)

Set-StrictMode -Version 3
$ErrorActionPreference = 'Stop'
$fmt = [StringComparison]::OrdinalIgnoreCase
$separators = [char[]]@([IO.Path]::DirectorySeparatorChar, [IO.Path]::AltDirectorySeparatorChar)
$storyTypes = 'document|header[0-9]+|footer[0-9]+|footnotes|endnotes|comments|glossary/document'
$dangerFields = '^(ASK|DATABASE|DDE|DDEAUTO|FILLIN|INCLUDEPICTURE|INCLUDETEXT|LINK|RD)$'
$refErrors = '(?i)Error\s*!\s*(Reference source not found|Bookmark not defined|No text of specified style)|(?:错误\s*[!！]\s*)?(?:未找到引用源|找不到引用源|引用源不存在|未定义书签|未找到指定样式|没有指定样式)'

function Get-FullLocalPath([string]$Path) {
    $clean = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
    if ($clean -notmatch '^[A-Za-z]:[\\/]') { throw "Expected a local drive path: $Path" }
    $clean = [IO.Path]::TrimEndingDirectorySeparator([IO.Path]::GetFullPath($clean))
    foreach ($component in $clean.Substring(3).Split($separators, [StringSplitOptions]::RemoveEmptyEntries)) {
        if ($component.EndsWith('.') -or $component.EndsWith(' ') -or $component.Contains(':') -or
            $component -match '^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\.|$)') {
            throw "Unsafe Windows path component: $component"
        }
    }
    return $clean
}

function Assert-PlainPath([string]$Path) {
    for ($p = $Path; $p; $p = [IO.Path]::GetDirectoryName($p)) {
        if (Test-Path -LiteralPath $p) {
            $actual = Get-Item -LiteralPath $p -Force
            if (($actual.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
                throw "Reparse points and symbolic links are not supported: $p"
            }
        }
    }
}

function Is-Within([string]$Child, [string]$ParentPath) {
    $prefix = $ParentPath.TrimEnd($separators) + [IO.Path]::DirectorySeparatorChar
    return $Child.Equals($ParentPath, $fmt) -or $Child.StartsWith($prefix, $fmt)
}

function Read-Part($Entry) {
    if ($Entry.Length -gt 96MB) { throw "Oversize XML part: $($Entry.FullName)" }
    $stream = $Entry.Open()
    $rd = $null
    try {
        $settings = [Xml.XmlReaderSettings]::new()
        $settings.DtdProcessing = [Xml.DtdProcessing]::Prohibit
        $settings.XmlResolver = $null
        $settings.MaxCharactersInDocument = 64MB
        $rd = [Xml.XmlReader]::Create($stream, $settings)
        $docXml = [Xml.XmlDocument]::new()
        $docXml.XmlResolver = $null
        $docXml.Load($rd)
        return ,$docXml
    } finally {
        if ($rd) { $rd.Dispose() }
        $stream.Dispose()
    }
}

function Read-FieldCounts([string]$Path) {
    $seq = [Collections.Generic.List[string]]::new()
    $pending = [Collections.Generic.List[object]]::new()
    $captionCount = 0
    $nRefErrors = 0
    $fieldCounts = @{}
    $archive = [IO.Compression.ZipFile]::OpenRead($Path)
    try {
        if (-not $archive.GetEntry('word/document.xml')) { throw "DOCX main part missing: $Path" }
        foreach ($entry in $archive.Entries) {
            if ($entry.FullName -notmatch '^(?:word/|_rels/|\[Content_Types\])') { continue }
            if ($entry.FullName -match '(?:^|/)vba(Project|Data)\.') { throw 'Macro parts are not supported.' }
            if ($entry.FullName -match '\.rels$') {
                $xml = Read-Part $entry
                foreach ($rel in $xml.DocumentElement.ChildNodes) {
                    if ($rel -is [Xml.XmlElement] -and $rel.GetAttribute('TargetMode') -eq 'External' -and $rel.GetAttribute('Type') -notmatch '/hyperlink$') {
                        throw "External document relationship cannot be updated safely: $($entry.FullName)"
                    }
                }
                continue
            }
            if ($entry.FullName -notmatch "^word/(?:$storyTypes)\.xml$") { continue }
            $xml = Read-Part $entry
            $ns = $xml.DocumentElement.NamespaceURI
            $stack = [Collections.Generic.List[object]]::new()
            $partFields = [Collections.Generic.List[object]]::new()
            $n = [Xml.XmlNamespaceManager]::new($xml.NameTable)
            $n.AddNamespace('w', $ns)
            foreach ($node in $xml.SelectNodes('//w:fldChar | //w:instrText | //w:fldSimple | //w:t', $n)) {
                switch ($node.LocalName) {
                    'fldSimple' {
                        $partFields.Add([pscustomobject]@{
                            code = $node.GetAttribute('instr', $ns)
                            result = (($node.SelectNodes('.//w:t', $n) | ForEach-Object { $_.InnerText }) -join '')
                        })
                    }
                    'fldChar' {
                        switch ($node.GetAttribute('fldCharType', $ns)) {
                            'begin' {
                                $state = [pscustomobject]@{ code = [Text.StringBuilder]::new(); result = [Text.StringBuilder]::new(); separate = $false }
                                $stack.Add($state)
                                $partFields.Add($state)
                            }
                            'separate' { if ($stack.Count) { $stack[$stack.Count - 1].separate = $true } }
                            'end' { if ($stack.Count) { $stack.RemoveAt($stack.Count - 1) } }
                        }
                    }
                    'instrText' {
                        if ($stack.Count -and -not $stack[$stack.Count - 1].separate) { [void]$stack[$stack.Count - 1].code.Append($node.InnerText) }
                    }
                    't' {
                        if ($stack.Count -and $stack[$stack.Count - 1].separate) { [void]$stack[$stack.Count - 1].result.Append($node.InnerText) }
                    }
                }
            }
            if ($stack.Count) { throw "Unclosed complex Word field: $($entry.FullName)" }
            foreach ($f in $partFields) { $pending.Add([pscustomobject]@{pt = $entry.FullName; f = $f}) }
        }
    } finally { $archive.Dispose() }

    # Sort parts for determinism; retain field order within each part.
    foreach ($record in ($pending | Sort-Object -Stable -Property pt)) {
        $f = $record.f
        if ([string]$f.code -notmatch '^\s*([A-Za-z]+)\b') { continue }
        $k = $Matches[1].ToUpperInvariant()
        if ($k -match $dangerFields) { throw "Field $k may load external content or prompt for input: $($record.pt)" }
        if (-not $fieldCounts.ContainsKey($k)) { $fieldCounts[$k] = 0 }
        $fieldCounts[$k]++
        if ($k -eq 'SEQ' -and [string]$f.code -match '^\s*SEQ\s+(?:"(?<label>[^"]+)"|(?<label>[^\s\\]+))') {
            $label = $Matches['label']
            $key = $record.pt + ':' + $label
            $seq.Add($key)
            if ($label -in $CaptionLabels -and [string]$f.code -notmatch '\\c\b') { ++$captionCount }
        }
        if ($k -in @('REF','PAGEREF','STYLEREF','NOTEREF','SEQ') -and [string]$f.result -match $refErrors) { ++$nRefErrors }
    }
    return [ordered]@{
        seq_fields = [int]$fieldCounts['SEQ']; caption_sequence_fields = $captionCount;
        ref_fields = [int]$fieldCounts['REF']; pageref_fields = [int]$fieldCounts['PAGEREF']; styleref_fields = [int]$fieldCounts['STYLEREF'];
        toc_fields = [int]$fieldCounts['TOC']; all_fields = $fieldCounts; sequence_order = @($seq.ToArray()); reference_errors = $nRefErrors
    }
}

try {
    if (-not $IsWindows) { throw 'This helper requires Windows (PowerShell 7+).' }
    $out = Get-FullLocalPath $OutputDirectory
    Assert-PlainPath $out
    if (Test-Path -LiteralPath $out) { throw "Output directory must not already exist: $out" }
    $names = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    $reports = [Collections.Generic.List[object]]::new()
    foreach ($inputFile in $Inputs) {
        $src = Get-FullLocalPath $inputFile
        Assert-PlainPath $src
        if ([IO.Path]::GetExtension($src) -ine '.docx' -or -not (Test-Path -LiteralPath $src -PathType Leaf)) { throw "Input must be an existing .docx file: $src" }
        $parent = [IO.Path]::GetDirectoryName($src)
        if ((Is-Within $out $parent) -or (Is-Within $parent $out)) { throw "Output directory overlaps an input's directory: $parent" }
        $name = [IO.Path]::GetFileName($src)
        if (-not $names.Add($name)) { throw "Duplicate report filename: $name" }
        $hash = (Get-FileHash -LiteralPath $src -Algorithm SHA256).Hash.ToLowerInvariant()
        $structure = Read-FieldCounts $src
        $reports.Add([ordered]@{
            source=$src; verified_copy=(Join-Path $out $name); source_sha256=$hash; output_sha256=$null;
            source_unchanged=$null; status='pending'; input_structure=$structure; saved_structure=$null;
            story_field_updates=@(); toc_updated=$null; native_toc_count=$null; native_main_story_fields=$null;
            saved_reopened=$false; page_count=$null; errors=@()
        })
    }
    if ($PreflightOnly) {
        [ordered]@{ mode='preflight_only'; write_performed=$false; report_count=$reports.Count; reports=$reports.ToArray() } | ConvertTo-Json -Depth 10
        exit 0
    }
    if (Test-Path -LiteralPath $out) { throw "Output already exists: $out" }
    [void](New-Item -ItemType Directory -Path $out)
    $logPath = Join-Path $out 'word-validation.json'
} catch {
    [Console]::Error.WriteLine($_.Exception.Message)
    exit 2
}

$history = [ordered]@{
    mode='native_word'; status='running'; started_utc=[DateTime]::UtcNow.ToString('o'); finished_utc=$null;
    word_version=$null; word_build=$null; exit_code=$null; error=$null; reports=$reports.ToArray()
}
function Write-RunLog {
    $temp = Join-Path $out ('.word-verification-' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $payload = $history | ConvertTo-Json -Depth 12
    [IO.File]::WriteAllText($temp, $payload + "`n", [Text.UTF8Encoding]::new($false))
    [IO.File]::Move($temp, $logPath, $true)
}

$word = $null
$doc = $null
$ownedSession = $false
$failed = $false
try {
    Write-RunLog
    $word = New-Object -ComObject Word.Application
    if ($word.Documents.Count -gt 0 -or $word.ProtectedViewWindows.Count -gt 0 -or $word.Visible) {
        throw 'Word returned an existing interactive session; no documents were touched.'
    }
    $ownedSession = $true
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $word.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
    $word.ScreenUpdating = $false
    $history.word_version = [string]$word.Version
    try { $history.word_build = [string]$word.Build } catch { $history.word_build = $null }

    foreach ($r in $reports) {
        try {
            $r.status = 'copying'
            Write-RunLog
            [IO.File]::Copy($r.source, $r.verified_copy, $false)
            $copiedHash = (Get-FileHash -LiteralPath $r.verified_copy -Algorithm SHA256).Hash.ToLowerInvariant()
            if ($copiedHash -ne $r.source_sha256) { throw 'Input changed before copying.' }
            $fileAttr = [IO.File]::GetAttributes($r.verified_copy)
            [IO.File]::SetAttributes($r.verified_copy, ($fileAttr -band (-bnot [IO.FileAttributes]::ReadOnly)))
            $r.status = 'updating_fields'
            Write-RunLog
            $doc = $word.Documents.Open([string]$r.verified_copy, $false, $false, $false)
            if ($doc.ReadOnly) { throw 'Verification copy was opened read-only and cannot be saved.' }
            foreach ($story in $doc.StoryRanges) {
                $range = $story
                while ($null -ne $range) {
                    $updateCode = $range.Fields.Update()
                    $r.story_field_updates += [ordered]@{ story_type=[int]$range.StoryType; word_return=$updateCode }
                    $range = $range.NextStoryRange
                }
            }
            $r.toc_updated = [int]$doc.TablesOfContents.Count
            foreach ($toc in $doc.TablesOfContents) { [void]$toc.Update() }
            foreach ($table in $doc.TablesOfFigures) { [void]$table.Update() }
            [void]$doc.Repaginate()
            foreach ($toc in $doc.TablesOfContents) { [void]$toc.UpdatePageNumbers() }
            foreach ($table in $doc.TablesOfFigures) { [void]$table.UpdatePageNumbers() }
            $doc.Save()
            $doc.Close(0)
            $doc = $null

            $r.status = 'reopening'
            Write-RunLog
            $doc = $word.Documents.Open([string]$r.verified_copy, $false, $true, $false)
            $r.saved_reopened = $true
            $r.page_count = [int]$doc.ComputeStatistics(2)
            $r.native_main_story_fields = [int]$doc.Fields.Count
            $r.native_toc_count = [int]$doc.TablesOfContents.Count
            $doc.Close(0)
            $doc = $null
            $r.saved_structure = Read-FieldCounts $r.verified_copy
            $r.output_sha256 = (Get-FileHash -LiteralPath $r.verified_copy -Algorithm SHA256).Hash.ToLowerInvariant()

            if (($r.input_structure.sequence_order -join "`n") -cne ($r.saved_structure.sequence_order -join "`n")) {
                $r.errors += 'SEQ type/order changed on Word save.'
            }
            if ($r.saved_structure.reference_errors -ne 0) { $r.errors += "Saved reference error fields: $($r.saved_structure.reference_errors)." }
            if ($r.native_toc_count -ne $r.toc_updated) { $r.errors += 'Native TOC count changed on reopening.' }
            if ($r.input_structure.toc_fields -ne $r.saved_structure.toc_fields) { $r.errors += 'TOC field count changed on Word save.' }
            if ($r.page_count -lt 1) { $r.errors += 'Word did not return a positive page count.' }
        } catch {
            $r.errors += $_.Exception.Message
        } finally {
            if ($doc) {
                try { $doc.Close(0) } catch { $r.errors += $_.Exception.Message }
                $doc = $null
            }
            try { $r.source_unchanged = (Get-FileHash -LiteralPath $r.source -Algorithm SHA256).Hash -ieq $r.source_sha256 } catch { $r.source_unchanged = $false }
            if (-not $r.source_unchanged) { $r.errors += 'Input hash changed or could not be rechecked.' }
            if ($r.errors.Count) { $r.status = 'failed'; $failed = $true } else { $r.status = 'passed' }
            Write-RunLog
            Write-Output ("{0}: {1}" -f [IO.Path]::GetFileName($r.source), $r.status)
        }
    }
} catch {
    $failed = $true
    $history.error = $_.Exception.Message
    [Console]::Error.WriteLine($history.error)
} finally {
    if ($word -and $ownedSession) {
        try { $word.Quit(0) } catch {
            $failed = $true
            $history.error = 'Word could not close the dedicated session: ' + $_.Exception.Message
        }
    }
    $history.finished_utc = [DateTime]::UtcNow.ToString('o')
    if ($failed) { $history.status = 'failed'; $history.exit_code = 1 } else { $history.status = 'passed'; $history.exit_code = 0 }
    Write-RunLog
}
Write-Output "Native Word log: $logPath"
exit $history.exit_code
