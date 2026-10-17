 # Thinka Excel Converter v1.5 development | Windows PowerShell 5.1 | Direct Open XML (no Excel COM)
# 11 original Thinka columns + one optional, Excel-only "配件组号" column.
# The raw CSV parser/writer never converts KNX addresses to dates or numeric types.
param([switch]$SelfTest,[switch]$GuiSmokeTest)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.IO.Compression
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.Windows.Forms.Application]::EnableVisualStyles()
$script:rawHeaders = @('Room *','Accessory Id','Service Name *','Service Type *','Characteristic or Parameter','Characteristic or Parameter name *','read address','write address','manufacturer','model','serial number')
$script:cnHeaders = @('房间 *','配件 ID','服务名称 *','服务类型 *','特征或参数','特征或参数名称 *','读取组地址','写入组地址','制造商','型号','序列号')
$script:groupHeader = '配件组号'
$script:orderHeader = '原始行序'
$script:settingsPath = Join-Path $env:APPDATA 'ThinkaCsvConverter\output-folder.txt'
$desktop = [Environment]::GetFolderPath('Desktop')
$defaultFolder = Join-Path $desktop 'Thinka文件'
$script:outputBase = $defaultFolder
if ([IO.File]::Exists($script:settingsPath)) {
    try { $saved = [IO.File]::ReadAllText($script:settingsPath).Trim(); if ($saved -and [IO.Directory]::Exists($saved) -and $saved.TrimEnd('\') -ine $desktop.TrimEnd('\')) { $script:outputBase = $saved } } catch {}
}
[void][IO.Directory]::CreateDirectory($script:outputBase)

function New-ResultRows { return ,(New-Object 'System.Collections.Generic.List[object]') }
function New-Record([string[]]$values, [string]$group, [int]$line) {
    return [pscustomobject]@{ Values=$values; Group=$group; Line=$line; OriginOrder=($line-1) }
}
function Unique-Target([string]$folder,[string]$name,[string]$extension) {
    [void][IO.Directory]::CreateDirectory($folder)
    $candidate = Join-Path $folder ($name + $extension); $n=1
    while ([IO.File]::Exists($candidate)) { $candidate=Join-Path $folder ('{0}_{1}{2}' -f $name,$n,$extension); $n++ }
    return $candidate
}
function Get-CsvData([string]$path) {
    $utf8=New-Object System.Text.UTF8Encoding($false,$true)
    $parser=New-Object Microsoft.VisualBasic.FileIO.TextFieldParser($path,$utf8,$true)
    $parser.TextFieldType=[Microsoft.VisualBasic.FileIO.FieldType]::Delimited
    $parser.SetDelimiters(',');$parser.HasFieldsEnclosedInQuotes=$true
    $result=New-ResultRows; $line=0
    try {
        while(-not $parser.EndOfData) {
            $parts=$parser.ReadFields(); $line++
            if ($parts.Length -ne 11) { throw "CSV 第 $line 条记录不是 11 列（实际 $($parts.Length) 列），已停止。" }
            if ($line -eq 1) {
                for($c=0;$c -lt 11;$c++) { if ($parts[$c] -cne $script:rawHeaders[$c]) { throw "CSV 第 $($c+1) 列英文表头与 Thinka 原格式不符：$($parts[$c])" } }
            } else {
                [void]$result.Add((New-Record ([string[]]$parts) '' $line))
            }
        }
    } finally { $parser.Close() }
    if ($line -eq 0) { throw '空 CSV 文件。' }
    return ,$result
}
# XLSX is an Open XML ZIP package. Read it directly: no Excel process or per-cell COM calls.
$script:xlsxNs='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$script:relsNs='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
function Get-XlsxNs($doc) {
    $mgr=[System.Xml.XmlNamespaceManager]::new($doc.NameTable)
    $mgr.AddNamespace('x',$script:xlsxNs)
    # XmlNamespaceManager implements IEnumerable; PowerShell must return it as one object.
    return ,$mgr
}
function Read-ZipXml($zip,[string]$name) {
    $entry=$zip.GetEntry($name)
    if($null -eq $entry) { throw "Excel 文件缺少必要组件：$name" }
    $stream=$entry.Open()
    try {
        $doc=New-Object System.Xml.XmlDocument
        $doc.XmlResolver=$null
        $doc.PreserveWhitespace=$true
        $doc.Load($stream)
        # Keep the XmlDocument itself intact rather than enumerating XML child nodes.
        return ,$doc
    } finally { $stream.Dispose() }
}
function Get-SheetPath($zip) {
    $workbook=Read-ZipXml $zip 'xl/workbook.xml'
    [System.Xml.XmlNamespaceManager]$mgr=Get-XlsxNs $workbook
    $first=$workbook.SelectSingleNode('/x:workbook/x:sheets/x:sheet[1]',$mgr)
    if($null -eq $first) { throw 'Excel 工作簿中找不到工作表。' }
    $relId=$first.GetAttribute('id',$script:relsNs)
    $rels=Read-ZipXml $zip 'xl/_rels/workbook.xml.rels'
    $target=''
    foreach($rel in $rels.DocumentElement.ChildNodes) {
        if($rel.LocalName -eq 'Relationship' -and $rel.GetAttribute('Id') -ceq $relId) {
            $target=$rel.GetAttribute('Target'); break
        }
    }
    if(-not $target) { throw '无法定位 Excel 第一个工作表。' }
    $resolved=[uri]::new([uri]'https://xlsx.local/xl/workbook.xml',$target.Replace([char]92,[char]47))
    $path=[uri]::UnescapeDataString($resolved.AbsolutePath).TrimStart('/')
    if(-not $path.StartsWith('xl/')) { throw 'Excel 工作表路径异常。' }
    return $path
}
function Get-ColumnNumber([string]$reference) {
    $number=0
    foreach($ch in $reference.ToUpperInvariant().ToCharArray()) {
        if($ch -lt 'A' -or $ch -gt 'Z') {break}
        $number=$number*26+([int][char]$ch-[int][char]'A'+1)
    }
    return $number
}
function Get-ColumnName([int]$number) {
    $name=''
    while($number -gt 0) {
        $number--; $name=([char](65+($number%26))).ToString()+$name
        $number=[int][math]::Floor($number/26)
    }
    return $name
}
function Get-SharedStrings($zip) {
    $values=New-Object 'System.Collections.Generic.List[string]'
    if($null -eq $zip.GetEntry('xl/sharedStrings.xml')) { return ,$values }
    $doc=Read-ZipXml $zip 'xl/sharedStrings.xml'
    [System.Xml.XmlNamespaceManager]$mgr=Get-XlsxNs $doc
    foreach($si in $doc.SelectNodes('/x:sst/x:si',$mgr)) {
        $builder=New-Object System.Text.StringBuilder
        foreach($t in $si.SelectNodes('x:t | x:r/x:t',$mgr)) { [void]$builder.Append($t.InnerText) }
        [void]$values.Add($builder.ToString())
    }
    return ,$values
}
function Read-CellText($cell,$ns,$strings) {
    if($null -eq $cell) {return ''}
    if($null -ne $cell.SelectSingleNode('x:f',$ns)) { throw "Excel 单元格 $($cell.GetAttribute('r')) 包含公式；请粘贴为普通文本。" }
    $kind=$cell.GetAttribute('t')
    if($kind -eq 'inlineStr') {
        $builder=New-Object System.Text.StringBuilder
        foreach($t in $cell.SelectNodes('x:is/x:t | x:is/x:r/x:t',$ns)) { [void]$builder.Append($t.InnerText) }
        return $builder.ToString()
    }
    $v=$cell.SelectSingleNode('x:v',$ns)
    if($null -eq $v) {return ''}
    if($kind -eq 's') {
        $i=0
        if(-not [int]::TryParse($v.InnerText,[ref]$i) -or $i -lt 0 -or $i -ge $strings.Count) { throw "无效的 Excel 共享文本索引：$($v.InnerText)" }
        return [string]$strings[$i]
    }
    if($kind -eq 'e') { throw "Excel 单元格 $($cell.GetAttribute('r')) 有错误值：$($v.InnerText)" }
    return [string]$v.InnerText
}
function Read-Xlsx([string]$path) {
    $zip=[IO.Compression.ZipFile]::OpenRead($path)
    try {
        $sheetName=Get-SheetPath $zip
        $doc=Read-ZipXml $zip $sheetName
        [System.Xml.XmlNamespaceManager]$ns=Get-XlsxNs $doc
        $strings=Get-SharedStrings $zip
        $rowNodes=$doc.SelectNodes('/x:worksheet/x:sheetData/x:row',$ns)
        if($rowNodes.Count -eq 0) { throw 'Excel 工作表为空。' }
        $columns=@{}
        $seen=@{}
        foreach($cell in $rowNodes[0].SelectNodes('x:c',$ns)) {
            $index=Get-ColumnNumber ($cell.GetAttribute('r'))
            $name=(Read-CellText $cell $ns $strings).Trim()
            if(-not $name) {continue}
            if($seen.ContainsKey($name)) {throw "Excel 表头【$name】重复，请检查第 1 行。"}
            $seen[$name]=$index
            $columns[$index]=$name
        }
        foreach($head in $script:cnHeaders) {
            if(-not $seen.ContainsKey($head)) {throw "Excel 缺少必需列【$head】，无法可靠转换。"}
        }
        $records=New-ResultRows
        foreach($row in $rowNodes) {
            $r=0
            if(-not [int]::TryParse($row.GetAttribute('r'),[ref]$r)) {continue}
            if($r -le 1) {continue}
            if($r -gt 100000) {throw 'Excel 使用区域超过10万行，请清理尾部空白格式。'}
            $values=@{}
            foreach($cell in $row.SelectNodes('x:c',$ns)) {
                $col=Get-ColumnNumber ($cell.GetAttribute('r'))
                if($columns.ContainsKey($col)) {
                    $name=$columns[$col]
                    if($name -in $script:cnHeaders -or $name -eq $script:groupHeader -or $name -eq $script:orderHeader) {
                        $values[$name]=Read-CellText $cell $ns $strings
                    }
                }
            }
            $v=[string[]]::new(11);$group='';$hasContent=$false
            if($seen.ContainsKey($script:groupHeader) -and $values.ContainsKey($script:groupHeader)) {
                $group=([string]$values[$script:groupHeader]).Trim()
            }
            for($c=0;$c -lt 11;$c++) {
                $head=$script:cnHeaders[$c]
                $v[$c]=if($values.ContainsKey($head)) {[string]$values[$head]} else {''}
                if($v[$c].Length -gt 0) {$hasContent=$true}
            }
            if($hasContent) {
                $record=New-Record $v $group $r
                if($values.ContainsKey($script:orderHeader)) {
                    $order=0
                    if([int]::TryParse(([string]$values[$script:orderHeader]),[ref]$order) -and $order -gt 0) {
                        $record.OriginOrder=$order
                    } else {$record.OriginOrder=0}
                } else {$record.OriginOrder=0}
                [void]$records.Add($record)
            }
        }
        return ,$records
    } finally {$zip.Dispose()}
}
function Get-ExportOrderedRows($records) {
    # Original rows keep their original relative order; Excel-added rows follow.
    # A copied '原始行序' is treated as a new row when that order has appeared before.
    $original=New-Object 'System.Collections.Generic.List[object]'
    $added=New-Object 'System.Collections.Generic.List[object]'
    $used=@{}
    foreach($record in $records) {
        $order=[int]$record.OriginOrder
        if($order -gt 0 -and -not $used.ContainsKey($order)) {
            $used[$order]=$true
            [void]$original.Add($record)
        } else {[void]$added.Add($record)}
    }
    $sorted=@($original | Sort-Object -Property OriginOrder)
    $result=New-ResultRows
    foreach($record in $sorted) {[void]$result.Add($record)}
    foreach($record in $added) {[void]$result.Add($record)}
    return ,$result
}
function Service-Key($values) { return ($values[0] + [char]31 + $values[2] + [char]31 + $values[3]) }
function Group-Key($values,[string]$group) { return ($values[0] + [char]31 + $group) }
function Unique-Strings($items) {
    $unique=New-Object 'System.Collections.Generic.List[string]'
    foreach($item in $items) { if($null -ne $item -and $item -ne '' -and -not $unique.Contains([string]$item)) { [void]$unique.Add([string]$item) } }
    return ,$unique.ToArray()
}
function Get-KnxGroupError([string]$value) {
    if($value -eq '') {return ''}
    $parts=$value.Split('/')
    if($parts.Length -eq 3) {
        if($value -cnotmatch '^\d+/\d+/\d+$') {return '不是有效的三段式 KNX 组地址'}
        $a=0;$b=0;$c=0
        if(-not [int]::TryParse($parts[0],[ref]$a) -or -not [int]::TryParse($parts[1],[ref]$b) -or -not [int]::TryParse($parts[2],[ref]$c)) {return '组地址数字过大'}
        if($a -gt 31 -or $b -gt 7 -or $c -gt 255) {return '三段式组地址超出 0-31/0-7/0-255 范围'}
        return ''
    }
    if($parts.Length -eq 2) {
        if($value -cnotmatch '^\d+/\d+$') {return '不是有效的两段式 KNX 组地址'}
        $a=0;$b=0
        if(-not [int]::TryParse($parts[0],[ref]$a) -or -not [int]::TryParse($parts[1],[ref]$b)) {return '组地址数字过大'}
        if($a -gt 31 -or $b -gt 2047) {return '两段式组地址超出 0-31/0-2047 范围'}
        return ''
    }
    return '组地址不是两段或三段形式（也可能被 Excel 转成日期）'
}
function Check-Records($records, [bool]$assignGrouped, [bool]$assignAll) {
    $errors=New-Object 'System.Collections.Generic.List[string]'
    $warnings=New-Object 'System.Collections.Generic.List[string]'
    $services=@{}; $groups=@{}; $fingerprints=@{}
    $uuidPattern='^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
    foreach($rec in $records) {
        $v=$rec.Values; $line=$rec.Line
        foreach($idx in @(0,2,3,4,5)) { if([string]::IsNullOrWhiteSpace($v[$idx])) { [void]$errors.Add("第 $line 行【$($script:cnHeaders[$idx])】不能为空。") } }
        if($v[1] -and $v[1] -cnotmatch $uuidPattern) { [void]$errors.Add("第 $line 行的配件 ID 不是标准 UUID 格式：$($v[1])") }
        if($v[4] -and $v[4] -cnotin @('characteristic','parameter')) { [void]$errors.Add("第 $line 行【特征或参数】应为 characteristic 或 parameter。") }
        if($v[4] -eq 'characteristic') {
            foreach($idx in @(6,7)) {
                if($v[$idx] -ne '') { $err=Get-KnxGroupError $v[$idx]; if($err) {[void]$errors.Add("第 $line 行【$($script:cnHeaders[$idx])】$($v[$idx])：$err") } }
            }
            if($v[6] -eq '' -and $v[7] -eq '') { [void]$warnings.Add("第 $line 行的特征 $($v[5]) 没有读写组地址，请确认是否有意留空。") }
        }
        if($v[4] -eq 'parameter' -and $v[7] -ne '') { [void]$warnings.Add("第 $line 行参数【$($v[5])】有写入地址，请确认 Thinka 是否支持该用法。") }
        $serviceKey=Service-Key $v
        if(-not $services.ContainsKey($serviceKey)) {
            $services[$serviceKey]=[pscustomobject]@{Key=$serviceKey;Room=$v[0];Name=$v[2];Type=$v[3];Members=(New-ResultRows);GroupLabels=(New-Object 'System.Collections.Generic.List[string]');IDs=(New-Object 'System.Collections.Generic.List[string]'); Group='';Id=''}
        }
        $svc=$services[$serviceKey]
        [void]$svc.Members.Add($rec)
        if($rec.Group -ne '' -and -not $svc.GroupLabels.Contains($rec.Group)) { [void]$svc.GroupLabels.Add($rec.Group) }
        if($v[1] -ne '' -and -not $svc.IDs.Contains($v[1])) { [void]$svc.IDs.Add($v[1]) }
        $fingerprint= $serviceKey + [char]31 + $v[4] + [char]31 + $v[5]
        if($fingerprints.ContainsKey($fingerprint)) { [void]$errors.Add("第 $line 行：同一服务的特征/参数【$($v[5])】重复，首次出现在第 $($fingerprints[$fingerprint]) 行。") }
        else {$fingerprints[$fingerprint]=$line}
    }
    foreach($svc in $services.Values) {
        if($svc.GroupLabels.Count -gt 1) { [void]$errors.Add("房间【$($svc.Room)】服务【$($svc.Name)】出现多个配件组号：$($svc.GroupLabels -join '、')") }
        if($svc.IDs.Count -gt 1) { [void]$errors.Add("房间【$($svc.Room)】服务【$($svc.Name)】的多行参数有不同配件 ID：$($svc.IDs -join '、')") }
        if($svc.GroupLabels.Count -eq 1) { $svc.Group=[string]$svc.GroupLabels[0] }
        if($svc.IDs.Count -eq 1) { $svc.Id=[string]$svc.IDs[0] }
        if($svc.Group -ne '') {
            $gkey=$svc.Room + [char]31 + $svc.Group
            if(-not $groups.ContainsKey($gkey)) { $groups[$gkey]=New-ResultRows }
            [void]$groups[$gkey].Add($svc)
        }
    }
    foreach($gkey in $groups.Keys) {
        $gsvc=$groups[$gkey]
        $differentIds=Unique-Strings @($gsvc | ForEach-Object { $_.Id })
        $blank= @($gsvc | Where-Object {$_.Id -eq ''}).Count
        if($differentIds.Length -eq 1 -and $blank -gt 0) {
            $first=$gsvc[0]
            [void]$warnings.Add("房间【$($first.Room)】配件组【$($first.Group)】同时包含已有 UUID 与空白 UUID；若由 Excel 复制设备产生，请核对组号，以免意外并入原配件。")
        }
        $sameType=@{}
        foreach($svc in $gsvc) {
            if($sameType.ContainsKey($svc.Type) -and $sameType[$svc.Type] -cne $svc.Name) {
                [void]$warnings.Add("配件组【$($svc.Group)】存在多个同类型服务【$($svc.Type)】；如果复制了设备，请检查配件组号。")
            } else {$sameType[$svc.Type]=$svc.Name}
        }
        if($differentIds.Length -gt 1) {
            $first=$gsvc[0]
            [void]$errors.Add("房间【$($first.Room)】组号【$($first.Group)】内已经存在不同配件 ID：$($differentIds -join '、')。请先手动确认，程序不会覆盖。")
        }
    }
    # Same UUID may not point to two differently typed services with identical name.
    $named=@{}
    foreach($svc in $services.Values) {
        if($svc.Id -eq '') {continue}
        $signature=$svc.Id.ToLowerInvariant()+[char]31+$svc.Name
        if($named.ContainsKey($signature) -and $named[$signature] -cne $svc.Type) {
            [void]$errors.Add("配件【$($svc.Id)】内服务名称【$($svc.Name)】重复，却具有不同服务类型：$($named[$signature]) / $($svc.Type)。")
        } else {$named[$signature]=$svc.Type}
    }
    # A copied old UUID can silently merge two separate devices. Flag suspicious same-type reuse.
    $idServiceTypes=@{}
    foreach($svc in $services.Values) {
        if($svc.Id -eq '') {continue}
        $tag=$svc.Id.ToLowerInvariant()+[char]31+$svc.Type
        if(-not $idServiceTypes.ContainsKey($tag)) {$idServiceTypes[$tag]=New-Object 'System.Collections.Generic.List[string]'}
        if(-not $idServiceTypes[$tag].Contains($svc.Name)) { [void]$idServiceTypes[$tag].Add($svc.Name) }
    }
    foreach($tag in $idServiceTypes.Keys) {
        if($idServiceTypes[$tag].Count -gt 1) {
            [void]$warnings.Add("一个 UUID 被多个同类型服务复用：$($idServiceTypes[$tag] -join '、')。如果是复制旧设备产生的，请清空新设备的旧 ID。")
        }
    }
    # If one accessory has services in different explicit groups, stop before modifying any ID.
    $ownerGroups=@{}
    foreach($svc in $services.Values) {
        if($svc.Id -eq '' -or $svc.Group -eq '') {continue}
        $owner=$svc.Id.ToLowerInvariant()
        $gkey=$svc.Room+[char]31+$svc.Group
        if($ownerGroups.ContainsKey($owner) -and $ownerGroups[$owner] -cne $gkey) {
            [void]$errors.Add("配件 ID【$owner】同时分配给了不同的显式配件组号，不能自动处理。")
        } else {$ownerGroups[$owner]=$gkey}
    }
    if($errors.Count -eq 0 -and ($assignGrouped -or $assignAll)) {
        foreach($gkey in $groups.Keys) {
            $gsvc=$groups[$gkey]
            $known=Unique-Strings @($gsvc | ForEach-Object { $_.Id })
            $chosen=if($known.Length -eq 1) { $known[0] } else { [Guid]::NewGuid().ToString() }
            foreach($svc in $gsvc) {$svc.Id=$chosen}
        }
        if($assignAll) {
            foreach($svc in $services.Values) { if($svc.Id -eq '') { $svc.Id=[Guid]::NewGuid().ToString() } }
        }
        # Before applying any changes check new effective IDs for conflicting service names.
        $named=@{}
        foreach($svc in $services.Values) {
            if($svc.Id -eq '') {continue}
            $signature=$svc.Id.ToLowerInvariant()+[char]31+$svc.Name
            if($named.ContainsKey($signature) -and $named[$signature] -cne $svc.Type) {
                [void]$errors.Add("分组后配件【$($svc.Id)】内服务名称【$($svc.Name)】冲突，类型：$($named[$signature]) / $($svc.Type)。请改服务名称。")
            } else { $named[$signature]=$svc.Type }
        }
    }
    $changed=0
    if($errors.Count -eq 0 -and ($assignGrouped -or $assignAll)) {
        foreach($svc in $services.Values) {
            if($svc.Id -eq '') {continue}
            foreach($rec in $svc.Members) {
                if($rec.Values[1] -ne $svc.Id) {$rec.Values[1]=$svc.Id;$changed++}
            }
        }
    }
    $allIds=Unique-Strings @($services.Values | ForEach-Object {$_.Id})
    return [pscustomobject]@{ Errors=$errors; Warnings=$warnings; Services=$services.Count; Groups=$groups.Count; IDs=$allIds.Length; Rows=$records.Count; UpdatedRows=$changed; Records=$records }
}
function Throw-IfInvalid($checked) {
    if($checked.Errors.Count -gt 0) {
        $first=@($checked.Errors | Select-Object -First 6)
        throw "检测到 $($checked.Errors.Count) 项错误，已停止，文件未输出：`r`n$($first -join "`r`n")"
    }
}
function Add-ZipText($zip,[string]$name,[string]$text) {
    $entry=$zip.CreateEntry($name,[IO.Compression.CompressionLevel]::Fastest)
    $stream=$entry.Open()
    try {
        $writer=[System.IO.StreamWriter]::new($stream,[System.Text.UTF8Encoding]::new($false))
        try {$writer.Write($text)} finally {$writer.Dispose()}
    } finally {$stream.Dispose()}
}
# Excel-compatible text-only styles. Numeric-looking GA/UUID/parameters remain text.
$script:stylesXml=@'
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2"><font><sz val="10"/><name val="Microsoft YaHei UI"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Microsoft YaHei UI"/></font></fonts>
  <fills count="5"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF194578"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FFF0F6FF"/><bgColor indexed="64"/></patternFill></fill><fill><patternFill patternType="solid"><fgColor rgb="FF51427D"/><bgColor indexed="64"/></patternFill></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="4"><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="49" fontId="1" fillId="2" borderId="0" xfId="0" applyFill="1" applyFont="1" applyNumberFormat="1"/><xf numFmtId="49" fontId="0" fillId="3" borderId="0" xfId="0" applyFill="1" applyNumberFormat="1"/><xf numFmtId="49" fontId="1" fillId="4" borderId="0" xfId="0" applyFill="1" applyFont="1" applyNumberFormat="1"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
'@
function Write-XlsxCell($writer,[string]$ref,[string]$value,[int]$style) {
    if($value -eq '') {return}
    $writer.WriteStartElement('c',$script:xlsxNs)
    $writer.WriteAttributeString('r',$ref)
    $writer.WriteAttributeString('s',[string]$style)
    $writer.WriteAttributeString('t','inlineStr')
    $writer.WriteStartElement('is',$script:xlsxNs)
    $writer.WriteStartElement('t',$script:xlsxNs)
    $writer.WriteAttributeString('xml','space','http://www.w3.org/XML/1998/namespace','preserve')
    $writer.WriteString($value)
    $writer.WriteEndElement();$writer.WriteEndElement();$writer.WriteEndElement()
}
function Write-Xlsx([string]$target,$records) {
    $zip=$null;$file=$null
    try {
        $file=[IO.File]::Create($target)
        $zip=[IO.Compression.ZipArchive]::new($file,[IO.Compression.ZipArchiveMode]::Create,$false)
        Add-ZipText $zip '[Content_Types].xml' '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>'
        Add-ZipText $zip '_rels/.rels' '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        Add-ZipText $zip 'xl/workbook.xml' '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><bookViews><workbookView/></bookViews><sheets><sheet name="Thinka设备" sheetId="1" r:id="rId1"/></sheets></workbook>'
        Add-ZipText $zip 'xl/_rels/workbook.xml.rels' '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'
        Add-ZipText $zip 'xl/styles.xml' ($script:stylesXml.Trim())
        $entry=$zip.CreateEntry('xl/worksheets/sheet1.xml',[IO.Compression.CompressionLevel]::Fastest)
        $stream=$entry.Open()
        try {
            $opts=New-Object System.Xml.XmlWriterSettings
            $opts.Encoding=[System.Text.UTF8Encoding]::new($false)
            $opts.Indent=$false
            $writer=[System.Xml.XmlWriter]::Create($stream,$opts)
            try {
                $ns=$script:xlsxNs
                $lastRow=[math]::Max(1,$records.Count+1)
                $writer.WriteStartDocument()
                $writer.WriteStartElement('worksheet',$ns)
                $writer.WriteStartElement('dimension',$ns);$writer.WriteAttributeString('ref',"A1:M$lastRow");$writer.WriteEndElement()
                $writer.WriteStartElement('sheetViews',$ns)
                $writer.WriteStartElement('sheetView',$ns);$writer.WriteAttributeString('workbookViewId','0')
                $writer.WriteStartElement('pane',$ns)
                $writer.WriteAttributeString('xSplit','1');$writer.WriteAttributeString('ySplit','1')
                $writer.WriteAttributeString('topLeftCell','B2');$writer.WriteAttributeString('activePane','bottomRight');$writer.WriteAttributeString('state','frozen')
                $writer.WriteEndElement();$writer.WriteEndElement();$writer.WriteEndElement()
                $writer.WriteStartElement('sheetFormatPr',$ns);$writer.WriteAttributeString('defaultRowHeight','17');$writer.WriteEndElement()
                $writer.WriteStartElement('cols',$ns)
                $widths=@(14,17,41,27,23,22,38,19,19,21,23,23,12)
                for($c=1;$c -le 13;$c++) {
                    $writer.WriteStartElement('col',$ns)
                    $writer.WriteAttributeString('min',[string]$c);$writer.WriteAttributeString('max',[string]$c)
                    $writer.WriteAttributeString('width',[string]$widths[$c-1]);$writer.WriteAttributeString('customWidth','1')
                    if($c -eq 13) {$writer.WriteAttributeString('hidden','1')}
                    $writer.WriteEndElement()
                }
                $writer.WriteEndElement()
                $writer.WriteStartElement('sheetData',$ns)
                $writer.WriteStartElement('row',$ns);$writer.WriteAttributeString('r','1');$writer.WriteAttributeString('ht','32');$writer.WriteAttributeString('customHeight','1')
                $heads=@($script:groupHeader)+@($script:cnHeaders)+@($script:orderHeader)
                for($c=1;$c -le 12;$c++) { $style=if($c -eq 1){3}else{1};Write-XlsxCell $writer ((Get-ColumnName $c)+'1') ([string]$heads[$c-1]) $style }
                $writer.WriteEndElement()
                $row=2
                foreach($rec in $records) {
                    $writer.WriteStartElement('row',$ns);$writer.WriteAttributeString('r',[string]$row)
                    Write-XlsxCell $writer ('A'+$row) ([string]$rec.Group) 2
                    for($c=0;$c -lt 11;$c++) { Write-XlsxCell $writer ((Get-ColumnName ($c+2))+$row) ([string]$rec.Values[$c]) 0 }
                    if([int]$rec.OriginOrder -gt 0) { Write-XlsxCell $writer ('M'+$row) ([string]$rec.OriginOrder) 0 }
                    $writer.WriteEndElement();$row++
                }
                $writer.WriteEndElement()
                $writer.WriteStartElement('autoFilter',$ns);$writer.WriteAttributeString('ref',"A1:L$lastRow");$writer.WriteEndElement()
                $writer.WriteEndElement();$writer.WriteEndDocument();$writer.Flush()
            } finally {$writer.Dispose()}
        } finally {$stream.Dispose()}
    } finally { if($null -ne $zip){$zip.Dispose()} elseif($null -ne $file){$file.Dispose()} }
}
function Csv-Quote([string]$s) { return '"' + $s.Replace('"','""') + '"' }
function Write-Csv([string]$path,$records) {
    $builder=New-Object System.Text.StringBuilder
    [void]$builder.Append((@($script:rawHeaders | ForEach-Object { Csv-Quote $_ }) -join ','))
    foreach($rec in $records) {
        [void]$builder.Append("`n")
        [void]$builder.Append((@($rec.Values | ForEach-Object { Csv-Quote $_ }) -join ','))
    }
    $utf8=New-Object System.Text.UTF8Encoding($false)
    [IO.File]::WriteAllText($path,$builder.ToString(),$utf8)
}
function Save-EditedExcel([string]$source,[string]$destination,$records) {
    # Patch only recognized Thinka columns, preserving all other XLSX ZIP entries,
    # other user-created columns, formatting, styles and row positioning.
    $sourceZip=$null;$destinationZip=$null;$destinationFile=$null
    try {
        $sourceZip=[IO.Compression.ZipFile]::OpenRead($source)
        $sheetPath=Get-SheetPath $sourceZip
        $doc=Read-ZipXml $sourceZip $sheetPath
        [System.Xml.XmlNamespaceManager]$ns=Get-XlsxNs $doc
        $strings=Get-SharedStrings $sourceZip
        $first=$doc.SelectSingleNode('/x:worksheet/x:sheetData/x:row[1]',$ns)
        if($null -eq $first) {throw '找不到 Excel 标题行。'}
        $mapping=@{}
        foreach($cell in $first.SelectNodes('x:c',$ns)) {
            $mapping[(Read-CellText $cell $ns $strings).Trim()]=Get-ColumnNumber ($cell.GetAttribute('r'))
        }
        foreach($head in $script:cnHeaders) {
            if(-not $mapping.ContainsKey($head)) {throw "Excel 缺少必要列【$head】。"}
        }
        if(-not $mapping.ContainsKey($script:groupHeader)) {
            foreach($rec in $records) {
                if($rec.Group -ne '') {throw 'Excel 缺少「配件组号」列，无法安全保存组合关系。请使用本工具生成的 Excel 模板。'}
            }
        }
        $rows=@{}
        foreach($node in $doc.SelectNodes('/x:worksheet/x:sheetData/x:row',$ns)) {$rows[[int]$node.GetAttribute('r')]=$node}
        foreach($rec in $records) {
            $r=[int]$rec.Line
            if(-not $rows.ContainsKey($r)) {throw "无法定位原 Excel 第 $r 行，请重新导入。"}
            $row=$rows[$r]
            $wanted=@{}
            for($c=0;$c -lt 11;$c++) {$wanted[$script:cnHeaders[$c]]=[string]$rec.Values[$c]}
            if($mapping.ContainsKey($script:groupHeader)) {$wanted[$script:groupHeader]=[string]$rec.Group}
            foreach($head in $wanted.Keys) {
                $col=$mapping[$head]
                $ref=(Get-ColumnName $col)+$r
                $cell=$row.SelectSingleNode("x:c[@r='$ref']",$ns)
                $before=Read-CellText $cell $ns $strings
                if($before -ceq $wanted[$head]) {continue}
                if($null -eq $cell) {
                    $cell=$doc.CreateElement('c',$script:xlsxNs)
                    $cell.SetAttribute('r',$ref)
                    $added=$false
                    foreach($next in $row.SelectNodes('x:c',$ns)) {
                        if((Get-ColumnNumber ($next.GetAttribute('r'))) -gt $col) {
                            [void]$row.InsertBefore($cell,$next);$added=$true;break
                        }
                    }
                    if(-not $added) {[void]$row.AppendChild($cell)}
                }
                foreach($child in @($cell.ChildNodes)) {[void]$cell.RemoveChild($child)}
                $cell.SetAttribute('t','inlineStr')
                $is=$doc.CreateElement('is',$script:xlsxNs)
                $t=$doc.CreateElement('t',$script:xlsxNs)
                $t.SetAttribute('space','http://www.w3.org/XML/1998/namespace','preserve')
                $t.InnerText=[string]$wanted[$head]
                [void]$is.AppendChild($t);[void]$cell.AppendChild($is)
            }
        }
        $destinationFile=[IO.File]::Create($destination)
        $destinationZip=[IO.Compression.ZipArchive]::new($destinationFile,[IO.Compression.ZipArchiveMode]::Create,$false)
        foreach($entry in $sourceZip.Entries) {
            $copy=$destinationZip.CreateEntry($entry.FullName,[IO.Compression.CompressionLevel]::Fastest)
            $out=$copy.Open()
            try {
                if($entry.FullName -ceq $sheetPath) {$doc.Save($out)}
                else {
                    $input=$entry.Open()
                    try {$input.CopyTo($out)} finally {$input.Dispose()}
                }
            } finally {$out.Dispose()}
        }
    } finally {
        if($null -ne $destinationZip) {$destinationZip.Dispose()}
        elseif($null -ne $destinationFile) {$destinationFile.Dispose()}
        if($null -ne $sourceZip) {$sourceZip.Dispose()}
    }
}
function Assert-ExportEqual($expected,$actual,[bool]$checkGroup) {
    if($expected.Count -ne $actual.Count) {throw "导出回读行数不一致：$($expected.Count) / $($actual.Count)。"}
    for($i=0;$i -lt $expected.Count;$i++) {
        for($j=0;$j -lt 11;$j++) {
            if($expected[$i].Values[$j] -cne $actual[$i].Values[$j]) {
                throw "导出回读发现第 $($i+1) 条记录【$($script:cnHeaders[$j])】不同，请检查数据。"
            }
        }
        if($checkGroup -and $expected[$i].Group -cne $actual[$i].Group) {
            throw "导出回读发现第 $($i+1) 条记录配件组号不同。"
        }
    }
}
function Clone-Records($records) {
    $copy=New-ResultRows
    foreach($rec in $records) {
        $item=New-Record ([string[]]@($rec.Values)) ([string]$rec.Group) ([int]$rec.Line)
        $item.OriginOrder=[int]$rec.OriginOrder
        [void]$copy.Add($item)
    }
    return ,$copy
}
function Get-ProjectStem([string]$path) {
    $stem=[IO.Path]::GetFileNameWithoutExtension($path)
    return ($stem -replace '_\d{8}_\d{4,6}(?:_\d+)?$','')
}
function New-ExportName([string]$path) {
    return ((Get-ProjectStem $path)+'_'+(Get-Date -Format 'yyyyMMdd_HHmmss'))
}
function Export-SafeCsv([string]$projectPath,$records,[bool]$allIDs) {
    $work=Clone-Records (Get-ExportOrderedRows $records)
    $check=Check-Records $work $true $allIDs
    Throw-IfInvalid $check
    $folder=Join-Path $script:outputBase 'Thinka导入文件'
    [void][IO.Directory]::CreateDirectory($folder)
    $target=Unique-Target $folder (New-ExportName $projectPath) '.csv'
    $temp=Join-Path $folder ([guid]::NewGuid().ToString('N')+'.tmp')
    try {
        Write-Csv $temp $work
        $reloaded=Get-CsvData $temp
        Assert-ExportEqual $work $reloaded $false
        [IO.File]::Move($temp,$target)
        # Only after success do generated UUIDs enter the current working copy.
        foreach($row in $work) {
            foreach($existing in $records) {
                if([object]::ReferenceEquals($row,$existing)) {continue}
                if($row.Line -eq $existing.Line -and $row.OriginOrder -eq $existing.OriginOrder) {
                    $existing.Values[1]=$row.Values[1]
                    break
                }
            }
        }
        return [pscustomobject]@{Path=$target;Check=$check}
    } finally {if([IO.File]::Exists($temp)) {[IO.File]::Delete($temp)}}
}
function Export-SafeExcel([string]$projectPath,[string]$sourceExcel,$records) {
    $folder=Join-Path $script:outputBase 'Excel配置文件'
    [void][IO.Directory]::CreateDirectory($folder)
    $target=Unique-Target $folder (New-ExportName $projectPath) '.xlsx'
    $temp=Join-Path $folder ([guid]::NewGuid().ToString('N')+'.tmp')
    try {
        if($sourceExcel -and [IO.File]::Exists($sourceExcel)) {
            Save-EditedExcel $sourceExcel $temp $records
        } else { Write-Xlsx $temp $records }
        $reloaded=Read-Xlsx $temp
        Assert-ExportEqual $records $reloaded $true
        [IO.File]::Move($temp,$target)
        return [pscustomobject]@{Path=$target;Check=(Check-Records (Clone-Records $records) $false $false)}
    } finally {if([IO.File]::Exists($temp)) {[IO.File]::Delete($temp)}}
}
function Generate-AllExcel([string]$path) {
    $records=Read-Xlsx $path
    $check=Check-Records $records $true $true
    Throw-IfInvalid $check
    $target=Unique-Target (Join-Path $script:outputBase 'Excel配置文件') (([IO.Path]::GetFileNameWithoutExtension($path))+'_已补齐UUID') '.xlsx'
    # Save a new copy of the user's workbook, preserving custom worksheet formatting.
    Save-EditedExcel $path $target $records
    return [pscustomobject]@{Path=$target;Check=$check}
}
function Convert-ToXlsx([string]$source) {
    $records=Get-CsvData $source
    $check=Check-Records $records $false $false
    $target=Unique-Target (Join-Path $script:outputBase '修改') ([IO.Path]::GetFileNameWithoutExtension($source)) '.xlsx'
    Write-Xlsx $target $records
    return [pscustomobject]@{Path=$target;Check=$check}
}
function Convert-ToCsv([string]$source,[bool]$generateAll) {
    $records=Read-Xlsx $source
    $check=Check-Records $records $true $generateAll
    Throw-IfInvalid $check
    $target=Unique-Target (Join-Path $script:outputBase 'Thinka导入文件') ([IO.Path]::GetFileNameWithoutExtension($source)) '.csv'
    Write-Csv $target $records
    $updatedXlsx=''
    if($check.UpdatedRows -gt 0) {
        # Preserve generated UUID values in an editable Excel copy so later exports stay stable.
        $updatedXlsx=Unique-Target (Join-Path $script:outputBase '修改') (([IO.Path]::GetFileNameWithoutExtension($source))+'_已写入UUID') '.xlsx'
        try { Save-EditedExcel $source $updatedXlsx $records }
        catch { $updatedXlsx=''; Write-Log '提醒：CSV 已生成，但保存已补齐 UUID 的 Excel 副本失败。以后重复转换原 Excel 会生成不同 UUID，请不要直接重复导入。' }
    }
    return [pscustomobject]@{Path=$target;Check=$check;UpdatedXlsx=$updatedXlsx}
}
function Check-File([string]$path) {
    $extension=[IO.Path]::GetExtension($path).ToLowerInvariant()
    if($extension -eq '.csv') {$records=Get-CsvData $path} elseif($extension -eq '.xlsx') {$records=Read-Xlsx $path} else {throw '只能检测 CSV 或 XLSX。'}
    # For validation, determine potential grouped collisions, but do not alter the original source.
    $check=Check-Records $records $true $false
    if($extension -eq '.csv') {
        $fs=[IO.File]::OpenRead($path)
        try {
            if($fs.Length -ge 3) {
                $b0=$fs.ReadByte(); $b1=$fs.ReadByte(); $b2=$fs.ReadByte()
                if($b0 -eq 239 -and $b1 -eq 187 -and $b2 -eq 191) {
                    [void]$check.Errors.Add('CSV 以 UTF-8 BOM 开头；当前 Thinka 原始样本不带 BOM。请使用右侧导出无 BOM 版本。')
                }
            }
        } finally { $fs.Close() }
    }
    $status=if($check.Errors.Count -eq 0) {'通过基础结构与已知规则检测'} else {'未通过'}
    $lines=New-Object 'System.Collections.Generic.List[string]'
    foreach($line in @('Thinka CSV / Excel 合规检测报告','',"来源文件：$path","检测结论：$status","数据记录：$($check.Rows)","服务数量：$($check.Services)","显式配件组：$($check.Groups)","错误：$($check.Errors.Count)","提醒：$($check.Warnings.Count)",'', '【错误】')) {[void]$lines.Add($line)}
    if($check.Errors.Count -eq 0) {[void]$lines.Add('无')} else {foreach($i in $check.Errors) {[void]$lines.Add('× '+$i)}}
    [void]$lines.Add('');[void]$lines.Add('【提醒】')
    if($check.Warnings.Count -eq 0) {[void]$lines.Add('无')} else {foreach($i in $check.Warnings) {[void]$lines.Add('! '+$i)}}
    [void]$lines.Add('');[void]$lines.Add('说明：本工具只依据 CSV 结构、格式和本项目验证过的规则进行静态检查。无法保证 Thinka 能接受所有服务类型、参数组合，必须由 Thinka 导入实测确认。')
    $target=Unique-Target (Join-Path $script:outputBase '检查报告') ([IO.Path]::GetFileNameWithoutExtension($path)+'_检查报告') '.txt'
    [IO.File]::WriteAllLines($target,$lines.ToArray(),(New-Object System.Text.UTF8Encoding($true)))
    return [pscustomobject]@{Path=$target;Check=$check}
}


# Offline regression test. Run with -SelfTest or double-click "运行自检.cmd".
if($SelfTest) {
    $watch=[Diagnostics.Stopwatch]::StartNew()
    $folder=Join-Path ([IO.Path]::GetTempPath()) ('Thinka_v14_selftest_'+[guid]::NewGuid().ToString('N'))
    [void][IO.Directory]::CreateDirectory($folder)
    $script:outputBase=$folder
    try {
        $source=Join-Path $PSScriptRoot '测试数据\三行示例.csv'
        if(-not [IO.File]::Exists($source)) { throw '自检示例 CSV 丢失。' }
        $original=Get-CsvData $source
        $converted=Convert-ToXlsx $source
        $excelRows=Read-Xlsx $converted.Path
        if($original.Count -ne $excelRows.Count) {throw 'CSV -> Excel 行数错误。'}
        for($i=0;$i -lt $original.Count;$i++) {
            for($j=0;$j -lt 11;$j++) {
                if($original[$i].Values[$j] -cne $excelRows[$i].Values[$j]) {throw "CSV -> Excel 内容变化：第 $i 行，第 $j 列。"}
            }
        }
        $restored=Convert-ToCsv $converted.Path $false
        $csvRows=Get-CsvData $restored.Path
        for($i=0;$i -lt $original.Count;$i++) {
            for($j=0;$j -lt 11;$j++) {
                if($original[$i].Values[$j] -cne $csvRows[$i].Values[$j]) {throw "Excel -> CSV 内容变化：第 $i 行，第 $j 列。"}
            }
        }
        $all=Generate-AllExcel $converted.Path
        $filled=Read-Xlsx $all.Path
        if(-not $filled[0].Values[1] -or $filled[0].Values[1] -cne $filled[1].Values[1] -or $filled[0].Values[1] -ceq $filled[2].Values[1]) {throw '独立 UUID 分配或同服务多行共用 UUID 失败。'}
        $excelRows[0].Group='1'
        $excelRows[2].Group='1'
        $groupedXlsx=Unique-Target (Join-Path $folder '修改') '组合服务验证' '.xlsx'
        Write-Xlsx $groupedXlsx $excelRows
        $groupedCsv=Convert-ToCsv $groupedXlsx $false
        $grouped=Get-CsvData $groupedCsv.Path
        if(-not $grouped[0].Values[1] -or $grouped[0].Values[1] -cne $grouped[1].Values[1] -or $grouped[0].Values[1] -cne $grouped[2].Values[1]) { throw '显式组号合并 UUID 失败。' }
        $checked=Check-File $groupedCsv.Path
        if($checked.Check.Errors.Count -gt 0) {throw '自检合规检查未通过。'}
        # v1.5: verified Excel/CSV export and bulk-added rows.
        $safeExcel=Export-SafeExcel $source '' $original
        $safeRows=Read-Xlsx $safeExcel.Path
        Assert-ExportEqual $original $safeRows $true
        $safeCsv=Export-SafeCsv $source $safeRows $false
        $roundtrip=Get-CsvData $safeCsv.Path
        Assert-ExportEqual $original $roundtrip $false
        $newValues=[string[]]@($safeRows[2].Values)
        $newValues[2]='风速新增设备'
        $newValues[6]='1/0/3'
        $newValues[1]=''
        $newRec=New-Record $newValues '' 101
        $newRec.OriginOrder=0
        [void]$safeRows.Add($newRec)
        $bulk=Export-SafeCsv $source $safeRows $false
        $bulkRead=Get-CsvData $bulk.Path
        if($bulkRead.Count -ne 4 -or $bulkRead[3].Values[2] -cne '风速新增设备') {
            throw 'Excel 批量新增记录导出错误。'
        }
        $badRows=Clone-Records $safeRows
        $badRows[0].Values[6]='99/99/99'
        $failed=$false
        try {[void](Export-SafeCsv $source $badRows $false)} catch {$failed=$true}
        if(-not $failed) {throw '危险组地址被允许导出。'}
        if($safeRows[0].Values[6] -cne '1/0/1') {throw '失败的导出修改了源工作副本。'}
        # v1.5: independent header ordering and extra Excel columns must not corrupt Thinka fields.
        $variant=Join-Path $folder '改列顺序及自定义列.xlsx'
        $srcZip=[IO.Compression.ZipFile]::OpenRead($safeExcel.Path)
        $destZip=$null
        try {
            $sheetPath=Get-SheetPath $srcZip
            $sheet=Read-ZipXml $srcZip $sheetPath
            [System.Xml.XmlNamespaceManager]$ns=Get-XlsxNs $sheet
            foreach($row in $sheet.SelectNodes('/x:worksheet/x:sheetData/x:row',$ns)) {
                $num=$row.GetAttribute('r')
                $b=$row.SelectSingleNode("x:c[@r='B$num']",$ns)
                $h=$row.SelectSingleNode("x:c[@r='H$num']",$ns)
                if($null -ne $b -and $null -ne $h) {
                    $b.SetAttribute('r',"H$num")
                    $h.SetAttribute('r',"B$num")
                }
                if($num -eq '1' -or $num -eq '2') {
                    $newCell=$sheet.CreateElement('c',$script:xlsxNs)
                    $newCell.SetAttribute('r',('N'+$num))
                    $newCell.SetAttribute('t','inlineStr')
                    $inner=$sheet.CreateElement('is',$script:xlsxNs)
                    $value=$sheet.CreateElement('t',$script:xlsxNs)
                    $value.InnerText=if($num -eq '1'){'备注'}else{'现场原始备注'}
                    [void]$inner.AppendChild($value)
                    [void]$newCell.AppendChild($inner)
                    [void]$row.AppendChild($newCell)
                }
            }
            $destZip=[IO.Compression.ZipArchive]::new([IO.File]::Create($variant),[IO.Compression.ZipArchiveMode]::Create,$false)
            foreach($entry in $srcZip.Entries) {
                $copy=$destZip.CreateEntry($entry.FullName,[IO.Compression.CompressionLevel]::Fastest)
                $out=$copy.Open()
                try {
                    if($entry.FullName -ceq $sheetPath) {$sheet.Save($out)}
                    else {
                        $input=$entry.Open()
                        try {$input.CopyTo($out)} finally {$input.Dispose()}
                    }
                } finally {$out.Dispose()}
            }
        } finally {
            if($null -ne $destZip) {$destZip.Dispose()}
            $srcZip.Dispose()
        }
        $columnsRead=Read-Xlsx $variant
        Assert-ExportEqual $original $columnsRead $true
        $columnsRead[0].Values[6]='1/0/7'
        $preserved=Export-SafeExcel $variant $variant $columnsRead
        $verify=Read-Xlsx $preserved.Path
        if($verify[0].Values[6] -cne '1/0/7') {throw '重排列后的编辑组地址没有被保存。'}
        $checkZip=[IO.Compression.ZipFile]::OpenRead($preserved.Path)
        try {
            $checkDoc=Read-ZipXml $checkZip (Get-SheetPath $checkZip)
            [System.Xml.XmlNamespaceManager]$checkNs=Get-XlsxNs $checkDoc
            $extra=$checkDoc.SelectSingleNode("//x:c[@r='N2']",$checkNs)
            if((Read-CellText $extra $checkNs (Get-SharedStrings $checkZip)) -cne '现场原始备注') {
                throw '保存 Excel 时丢失了自定义列。'
            }
        } finally {$checkZip.Dispose()}
        $watch.Stop()
        Write-Host ('自检通过：双向转换、批量新增、安全回读、UUID 分组、错误阻断和工作副本保护。') -ForegroundColor Green
        Write-Host ('自检共耗时：'+$watch.Elapsed.TotalSeconds.ToString('F2')+' 秒。测试文件：'+$folder)
        exit 0
    } catch {
        Write-Host ('自检失败：'+$_.Exception.Message) -ForegroundColor Red
        Write-Host ('调试文件位置：'+$folder)
        exit 1
    }
}

# v1.5 single-project user interface. Excel is the primary batch-entry surface.
$script:project=New-ResultRows
$script:baseline=New-ResultRows
$script:projectPath=''
$script:sourceExcel=''
$script:lastExcel=''
$script:lastCsv=''
$script:undo=New-Object 'System.Collections.Generic.Stack[object]'
$script:redo=New-Object 'System.Collections.Generic.Stack[object]'
$script:drawing=$false
$script:detailRefreshQueued=$false
$script:rebuildQueued=$false
$script:rebuildFocus=0
$script:visibleRecordByLine=@{}
function Record-Signature($rows) {
    $b=New-Object Text.StringBuilder
    foreach($r in $rows) {
        [void]$b.Append($r.Group)
        foreach($v in $r.Values) {[void]$b.Append([char]31);[void]$b.Append([string]$v)}
        [void]$b.Append([char]30)
    }
    return $b.ToString()
}
function Has-Unsaved {
    if(-not $script:projectPath) {return $false}
    $s=Record-Signature $script:project
    return ($s -cne $script:lastExcel -and $s -cne $script:lastCsv)
}
function Snapshot {
    $script:undo.Push((Clone-Records $script:project))
    $script:redo.Clear()
}
$form=New-Object Windows.Forms.Form
# On PS2EXE builds, use the icon embedded in the executable for the window.
try {
    $binary=[Diagnostics.Process]::GetCurrentProcess().MainModule.FileName
    if($binary -and [IO.File]::Exists($binary)) {
        $extracted=[Drawing.Icon]::ExtractAssociatedIcon($binary)
        if($null -ne $extracted) {$form.Icon=$extracted}
    }
} catch {}
$form.Text='Thinka Excel 转换工具 v1.5 开发测试版'
$form.Size=New-Object Drawing.Size(1150,730)
$form.MinimumSize=New-Object Drawing.Size(970,620)
$form.StartPosition='CenterScreen'
$form.Font=New-Object Drawing.Font('Microsoft YaHei UI',10)
$form.KeyPreview=$true
$tabs=New-Object Windows.Forms.TabControl
$tabs.Dock='Fill'
$form.Controls.Add($tabs)
$pageFiles=New-Object Windows.Forms.TabPage('文件操作')
$pagePreview=New-Object Windows.Forms.TabPage('设备预览与编辑')
$tabs.TabPages.AddRange([Windows.Forms.TabPage[]]@($pageFiles,$pagePreview))
$folderText=New-Object Windows.Forms.TextBox
$folderText.ReadOnly=$true
$folderText.Location=New-Object Drawing.Point(20,20)
$folderText.Width=880
$folderText.Anchor='Top,Left,Right'
$folderText.Text=$script:outputBase
$pageFiles.Controls.Add($folderText)
$browse=New-Object Windows.Forms.Button
$browse.Text='选择输出文件夹'
$browse.Location=New-Object Drawing.Point(915,19)
$browse.Size=New-Object Drawing.Size(170,30)
$browse.Anchor='Top,Right'
$pageFiles.Controls.Add($browse)
$browse.Add_Click({
    $d=New-Object Windows.Forms.FolderBrowserDialog
    $d.SelectedPath=$script:outputBase
    if($d.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) {
        $script:outputBase=$d.SelectedPath
        $folderText.Text=$script:outputBase
        [void][IO.Directory]::CreateDirectory((Split-Path $script:settingsPath))
        [IO.File]::WriteAllText($script:settingsPath,$script:outputBase)
    }
})
function File-Button([string]$title,[int]$x,[int]$y) {
    $b=New-Object Windows.Forms.Button
    $b.Text=$title
    $b.Location=New-Object Drawing.Point($x,$y)
    $b.Size=New-Object Drawing.Size(250,48)
    $pageFiles.Controls.Add($b)
    return $b
}
$loadCsv=File-Button '① 导入 Thinka CSV' 22 90
$loadXlsx=File-Button '② 导入编辑好的 Excel' 300 90
$saveXlsx=File-Button '③ 导出 Excel 配置文件' 22 161
$saveCsv=File-Button '④ 导出 Thinka CSV' 300 161
$reportBtn=File-Button '全面检测' 580 90
$previewBtn=File-Button '进入设备预览' 580 161
$generateAll=New-Object Windows.Forms.CheckBox
$generateAll.Text='导出 CSV 时，也为独立服务的空白 UUID 生成 ID（不勾选则交由 Thinka）'
$generateAll.Location=New-Object Drawing.Point(22,236)
$generateAll.Size=New-Object Drawing.Size(780,29)
$pageFiles.Controls.Add($generateAll)
$topmost=New-Object Windows.Forms.CheckBox
$topmost.Text='窗口置顶'
$topmost.Location=New-Object Drawing.Point(840,237)
$topmost.Size=New-Object Drawing.Size(120,28)
$pageFiles.Controls.Add($topmost)
$topmost.Add_CheckedChanged({$form.TopMost=$topmost.Checked})
$status=New-Object Windows.Forms.Label
$status.Text='请导入 Thinka CSV 或之前生成的 Excel。'
$status.Location=New-Object Drawing.Point(23,287)
$status.Size=New-Object Drawing.Size(1050,83)
$status.Anchor='Top,Left,Right'
$pageFiles.Controls.Add($status)
$log=New-Object Windows.Forms.TextBox
$log.Multiline=$true
$log.ReadOnly=$true
$log.ScrollBars='Vertical'
$log.Text='源文件不覆盖。推荐先导出 Excel，批量复制和编辑设备后重新导入。'
$log.Location=New-Object Drawing.Point(22,380)
$log.Size=New-Object Drawing.Size(1060,260)
$log.Anchor='Top,Bottom,Left,Right'
$pageFiles.Controls.Add($log)
function Write-Log([string]$message) {
    $log.AppendText([Environment]::NewLine+$message)
    $log.SelectionStart=$log.TextLength
    $log.ScrollToCaret()
}
$split=New-Object Windows.Forms.SplitContainer
$split.Dock='Fill'
$split.FixedPanel=[Windows.Forms.FixedPanel]::Panel1
$split.SplitterDistance=300
$pagePreview.Controls.Add($split)
# Assign splitter position after WinForms has sized the tab page. Otherwise
# the initial 194px Panel2 width can stay fixed while the form expands.
function Set-EditorPaneWidths {
    if($split.Width -lt 700) {return}
    $split.FixedPanel=[Windows.Forms.FixedPanel]::Panel1
    $target=[math]::Min(315,[math]::Max(235,$split.Width-510))
    if($split.SplitterDistance -ne $target) {$split.SplitterDistance=$target}
}
$split.Add_SizeChanged({Set-EditorPaneWidths})
$tabs.Add_SelectedIndexChanged({
    if($tabs.SelectedTab -eq $pagePreview) {
        Set-EditorPaneWidths
        $split.PerformLayout()
    }
})
$form.Add_Shown({Set-EditorPaneWidths})
$filter=New-Object Windows.Forms.TextBox
$filter.Dock='Top'
$split.Panel1.Controls.Add($filter)
$tree=New-Object Windows.Forms.TreeView
$tree.Dock='Fill'
$tree.HideSelection=$false
$split.Panel1.Controls.Add($tree)
$tree.BringToFront()
$filter.BringToFront()
# Explicit four-row layout prevents Dock=Fill DataGridView from rendering
# underneath the title/action bar (caused clipped header and first service row).
$rightLayout=New-Object Windows.Forms.TableLayoutPanel
$rightLayout.Dock='Fill'
$rightLayout.ColumnCount=1
$rightLayout.RowCount=4
$rightLayout.Margin=New-Object Drawing.Padding(0)
$rightLayout.Padding=New-Object Drawing.Padding(0)
[void]$rightLayout.ColumnStyles.Add(([Windows.Forms.ColumnStyle]::new([Windows.Forms.SizeType]::Percent,100)))
[void]$rightLayout.RowStyles.Add(([Windows.Forms.RowStyle]::new([Windows.Forms.SizeType]::Absolute,42)))
[void]$rightLayout.RowStyles.Add(([Windows.Forms.RowStyle]::new([Windows.Forms.SizeType]::Absolute,50)))
[void]$rightLayout.RowStyles.Add(([Windows.Forms.RowStyle]::new([Windows.Forms.SizeType]::Percent,100)))
[void]$rightLayout.RowStyles.Add(([Windows.Forms.RowStyle]::new([Windows.Forms.SizeType]::Absolute,132)))
$split.Panel2.Controls.Add($rightLayout)
$heading=New-Object Windows.Forms.Label
$heading.Dock='Fill'
$heading.TextAlign=[Drawing.ContentAlignment]::MiddleLeft
$heading.Margin=New-Object Drawing.Padding(8,0,4,0)
$heading.Text='从左侧选择设备'
[void]$rightLayout.Controls.Add($heading,0,0)
$buttons=New-Object Windows.Forms.FlowLayoutPanel
$buttons.Dock='Fill'
$buttons.Margin=New-Object Drawing.Padding(4,2,4,2)
$buttons.AutoScroll=$true
$buttons.WrapContents=$false
[void]$rightLayout.Controls.Add($buttons,0,1)
function Edit-Button([string]$title) {
    $b=New-Object Windows.Forms.Button
    $b.Text=$title
    $b.Size=New-Object Drawing.Size(102,31)
    [void]$buttons.Controls.Add($b)
    return $b
}
$undoBtn=Edit-Button '撤销 Ctrl+Z'
$redoBtn=Edit-Button '重做 Ctrl+Y'
$restoreBtn=Edit-Button '恢复此设备'
$groupBtn=Edit-Button '组合服务'
$separateBtn=Edit-Button '解除组合'
$advanced=New-Object Windows.Forms.CheckBox
$advanced.Text='高级编辑'
$advanced.Width=110
[void]$buttons.Controls.Add($advanced)
$grid=New-Object Windows.Forms.DataGridView
$grid.Dock='Fill'
$grid.Margin=New-Object Drawing.Padding(2,0,2,0)
$grid.AllowUserToResizeColumns=$true
$grid.AllowUserToResizeRows=$true
$grid.ScrollBars=[Windows.Forms.ScrollBars]::Both
$grid.RowHeadersVisible=$false
$grid.AllowUserToAddRows=$false
$grid.AllowUserToDeleteRows=$false
$grid.AutoSizeColumnsMode='None'
$grid.AutoGenerateColumns=$false
$grid.AutoSizeRowsMode=[Windows.Forms.DataGridViewAutoSizeRowsMode]::AllCells
$grid.DefaultCellStyle.WrapMode=[Windows.Forms.DataGridViewTriState]::True
$grid.DefaultCellStyle.Padding=New-Object Windows.Forms.Padding(2,3,2,3)
$grid.BackgroundColor=[Drawing.Color]::White
$grid.ColumnHeadersVisible=$true
$grid.ColumnHeadersHeightSizeMode=[Windows.Forms.DataGridViewColumnHeadersHeightSizeMode]::DisableResizing
$grid.ColumnHeadersHeight=43
$grid.ColumnHeadersDefaultCellStyle.Alignment=[Windows.Forms.DataGridViewContentAlignment]::MiddleLeft
$grid.ColumnHeadersDefaultCellStyle.WrapMode=[Windows.Forms.DataGridViewTriState]::False
$grid.EnableHeadersVisualStyles=$false
$grid.ColumnHeadersDefaultCellStyle.BackColor=[Drawing.Color]::FromArgb(232,239,248)
$grid.AlternatingRowsDefaultCellStyle.BackColor=[Drawing.Color]::FromArgb(247,250,253)
$grid.SelectionMode='FullRowSelect'
$grid.MultiSelect=$false
[void]$rightLayout.Controls.Add($grid,0,2)
$issueList=New-Object Windows.Forms.ListBox
$issueList.Dock='Fill'
$issueList.Margin=New-Object Drawing.Padding(4,3,4,3)
[void]$rightLayout.Controls.Add($issueList,0,3)
foreach($name in @('原始行号','房间','服务名称','服务类型','特征 / 参数','特征 / 参数名称','读取组地址','写入组地址','配件 UUID')) {
    [void]$grid.Columns.Add(('c'+$grid.Columns.Count),$name)
}
foreach($index in @(0,3,4,5)) {$grid.Columns[$index].ReadOnly=$true}
# Prefer readable complete values with a horizontal scrollbar instead of ellipses.
$columnWidths=@(65,115,210,205,190,310,155,155,295)
for($i=0;$i -lt $columnWidths.Length;$i++) {
    $grid.Columns[$i].DataPropertyName='Field'+$i
    $grid.Columns[$i].Width=$columnWidths[$i]
    $grid.Columns[$i].MinimumWidth=55
    $grid.Columns[$i].Resizable=[Windows.Forms.DataGridViewTriState]::True
    $grid.Columns[$i].DefaultCellStyle.WrapMode=[Windows.Forms.DataGridViewTriState]::True
}
$grid.ShowCellToolTips=$true
$grid.Add_CellFormatting({
    if($_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) {return}
    $cell=$grid.Rows[$_.RowIndex].Cells[$_.ColumnIndex]
    if($null -ne $cell.Value) {$cell.ToolTipText=[string]$cell.Value}
})
$grid.Add_DataError({
    Write-Log ("表格单元格错误：行 {0}，列 {1}；{2}" -f $_.RowIndex,$_.ColumnIndex,$_.Exception.Message)
    $_.ThrowException=$false
})
$grid.Columns[8].Visible=$false
function Accessory-Key($rec) {
    $v=$rec.Values
    $prefix=$v[0]+[char]31
    if($rec.Group -ne '') {return $prefix+'G'+$rec.Group}
    if($v[1] -ne '') {return $prefix+'I'+$v[1].ToLowerInvariant()}
    return $prefix+'S'+(Service-Key $v)
}
function Find-Line([int]$line) {
    foreach($room in $tree.Nodes) {
        foreach($device in $room.Nodes) {
            foreach($rec in $device.Tag) {
                if($rec.Line -eq $line) {return $device}
            }
        }
    }
    return $null
}
function Show-Device {
    # Data-bound snapshots avoid Rows.Clear()/Rows.Add() reentrancy while
    # DataGridView is committing a cell or changing the current row.
    if($script:drawing) {return}
    $script:drawing=$true
    try {
        $node=$tree.SelectedNode
        $dt=New-Object System.Data.DataTable
        for($c=0;$c -lt 9;$c++) {
            [void]$dt.Columns.Add(('Field'+$c),[string])
        }
        $lookup=@{}
        $shown=0
        if($null -ne $node -and $node.Level -eq 1) {
            foreach($rec in $node.Tag) {
                if($null -eq $rec -or $null -eq $rec.Values -or $rec.Values.Length -ne 11) {
                    throw '设备记录缺少标准 11 列字段，无法显示。'
                }
                $v=$rec.Values
                $line=[string]$rec.Line
                $cells=[object[]]@($line,$v[0],$v[2],$v[3],$v[4],$v[5],$v[6],$v[7],$v[1])
                [void]$dt.Rows.Add($cells)
                $lookup[$line]=$rec
                $shown++
            }
            $heading.Text=$node.Parent.Text+' / '+$node.Text+"    （$shown 行参数）"
        } else {
            $heading.Text='从左侧选择设备'
        }
        $script:visibleRecordByLine=$lookup
        $grid.DataSource=$dt
        if($grid.Rows.Count -gt 0) {
            try {$grid.FirstDisplayedScrollingRowIndex=0} catch {}
        }
        $grid.Refresh()
    } catch {
        $heading.Text='参数显示失败，请看文件操作页的日志'
        Write-Log ('设备参数显示失败：'+$_.Exception.Message)
        [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'设备参数显示错误')
    } finally {
        $script:drawing=$false
    }
}
function Queue-DetailRefresh {
    # Do not replace DataSource synchronously inside TreeView or grid callbacks.
    if($script:detailRefreshQueued) {return}
    if(-not $form.IsHandleCreated) {Show-Device;return}
    $script:detailRefreshQueued=$true
    [void]$form.BeginInvoke([System.Action]{
        try {Show-Device}
        finally {$script:detailRefreshQueued=$false}
    })
}
function Queue-Rebuild([int]$line) {
    # CellEndEdit / validation must return before changing the tree + grid.
    $script:rebuildFocus=$line
    if($script:rebuildQueued) {return}
    if(-not $form.IsHandleCreated) {Rebuild-Tree $line;return}
    $script:rebuildQueued=$true
    [void]$form.BeginInvoke([System.Action]{
        try {Rebuild-Tree $script:rebuildFocus}
        finally {$script:rebuildQueued=$false}
    })
}
function Show-Issues {
    $check=Check-Records (Clone-Records $script:project) $false $false
    $issueList.Items.Clear()
    foreach($err in $check.Errors) {[void]$issueList.Items.Add('错误：'+$err)}
    foreach($warning in $check.Warnings) {[void]$issueList.Items.Add('提醒：'+$warning)}
    $sign=Record-Signature $script:project
    $excel=if($sign -ceq $script:lastExcel){'最新版'}else{'未导出最新版'}
    $csv=if($sign -ceq $script:lastCsv){'最新版'}else{'未导出最新版'}
    $status.Text="项目：$([IO.Path]::GetFileName($script:projectPath))，记录 $($check.Rows)，服务 $($check.Services)，错误 $($check.Errors.Count)，提醒 $($check.Warnings.Count)" + [Environment]::NewLine + "Excel：$excel；CSV：$csv。"
    return $check
}
function Rebuild-Tree([int]$focusLine) {
    $script:drawing=$true
    try {
        $tree.BeginUpdate()
        $tree.Nodes.Clear()
        $rooms=@{}
        $items=@{}
        foreach($rec in $script:project) {
            $v=$rec.Values
            $all=($v -join ' ')+' '+$rec.Group
            if($filter.Text -and $all.IndexOf($filter.Text,[StringComparison]::OrdinalIgnoreCase) -lt 0) {continue}
            $room=[string]$v[0]
            $key=Accessory-Key $rec
            if(-not $rooms.ContainsKey($room)) {
                $rn=New-Object Windows.Forms.TreeNode($room)
                $rooms[$room]=$rn
                [void]$tree.Nodes.Add($rn)
            }
            if(-not $items.ContainsKey($key)) {
                $dn=New-Object Windows.Forms.TreeNode($v[2])
                $dn.Tag=New-ResultRows
                $items[$key]=$dn
                [void]$rooms[$room].Nodes.Add($dn)
            }
            [void]$items[$key].Tag.Add($rec)
        }
        foreach($device in $items.Values) {
            $names=Unique-Strings @($device.Tag | ForEach-Object {$_.Values[2]})
            if($names.Length -gt 1) {$device.Text=$names[0]+' ('+$names.Length+' 个服务)'}
        }
        $tree.ExpandAll()
        $tree.EndUpdate()
        $select=if($focusLine -gt 0){Find-Line $focusLine}else{$null}
        if($null -eq $select) {
            foreach($rn in $tree.Nodes) {if($rn.Nodes.Count -gt 0) {$select=$rn.Nodes[0];break}}
        }
        if($null -ne $select) {$tree.SelectedNode=$select}
    } finally {$script:drawing=$false}
    Queue-DetailRefresh
    [void](Show-Issues)
}
function Load-File([string]$path) {
    if(Has-Unsaved) {
        if([Windows.Forms.MessageBox]::Show('存在未导出修改，确认放弃并导入新文件？','切换项目','YesNo','Warning') -ne 'Yes') {return}
    }
    $ext=[IO.Path]::GetExtension($path).ToLowerInvariant()
    if($ext -eq '.csv') {$rows=Get-CsvData $path}
    elseif($ext -eq '.xlsx') {$rows=Read-Xlsx $path}
    else {throw '只接受 CSV 和 XLSX 文件。'}
    $script:project=$rows
    $script:baseline=Clone-Records $rows
    $script:projectPath=$path
    $script:sourceExcel=if($ext -eq '.xlsx'){$path}else{''}
    $script:lastCsv=if($ext -eq '.csv'){Record-Signature $rows}else{''}
    $script:lastExcel=if($ext -eq '.xlsx'){Record-Signature $rows}else{''}
    $script:undo.Clear();$script:redo.Clear()
    $filter.Text=''
    Rebuild-Tree 0
    Write-Log ('已导入：'+$path+'，'+$rows.Count+' 行。')
    $tabs.SelectedTab=$pagePreview
}
function Choose-Input([string]$kind) {
    $d=New-Object Windows.Forms.OpenFileDialog
    $d.Filter=if($kind -eq 'csv'){'Thinka CSV (*.csv)|*.csv'}else{'Excel (*.xlsx)|*.xlsx'}
    if($d.ShowDialog() -ne 'OK') {return}
    try {Load-File $d.FileName}
    catch {[void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'导入失败')}
}
function Do-Export([string]$kind) {
    if(-not $script:projectPath) {[void][Windows.Forms.MessageBox]::Show('请先导入文件。');return}
    try {
        $checked=Show-Issues
        if($checked.Errors.Count -gt 0) {
            [void][Windows.Forms.MessageBox]::Show("有 $($checked.Errors.Count) 项错误，先修复再导出。")
            return
        }
        if($kind -eq 'csv') {
            $msg="准备生成 Thinka CSV。" + [Environment]::NewLine + "记录 $($checked.Rows)，服务 $($checked.Services)，提醒 $($checked.Warnings.Count)。" + [Environment]::NewLine
            if($checked.Warnings.Count -gt 0) {
                $msg += '需要核对的提醒：' + [Environment]::NewLine
                $msg += ((@($checked.Warnings | Select-Object -First 5)) -join [Environment]::NewLine)
                $msg += [Environment]::NewLine
            }
            $msg += '确认导出？'
            if([Windows.Forms.MessageBox]::Show($msg,'确认导出','YesNo','Question') -ne 'Yes') {return}
            $res=Export-SafeCsv $script:projectPath $script:project ([bool]$generateAll.Checked)
            $script:lastCsv=Record-Signature $script:project
        } else {
            $res=Export-SafeExcel $script:projectPath $script:sourceExcel $script:project
            $script:lastExcel=Record-Signature $script:project
        }
        Rebuild-Tree 0
        Write-Log ('导出与回读校验成功：'+$res.Path)
        if([Windows.Forms.MessageBox]::Show("已导出："+[Environment]::NewLine+$res.Path+[Environment]::NewLine+'打开所在文件夹？','成功','YesNo','Information') -eq 'Yes') {
            [void][Diagnostics.Process]::Start('explorer.exe',('/select,"'+$res.Path+'"'))
        }
    } catch {
        Write-Log ('导出失败：'+$_.Exception.Message)
        [void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'导出失败')
    }
}
function Normalize-Address([string]$text) {
    $text=$text.Trim()
    if($text -match '^\d+\s+\d+\s+\d+$') {return ($text -replace '\s+','/')}
    if($text -match '^\d+\s*/\s*\d+(?:\s*/\s*\d+)?$') {return ($text -replace '\s*/\s*','/')}
    return $text
}
$grid.Add_CellEndEdit({
    if($script:drawing -or $_.RowIndex -lt 0 -or $_.ColumnIndex -lt 0) {return}
    $idx=$_.ColumnIndex;$row=$grid.Rows[$_.RowIndex]
    $line=[string]$row.Cells[0].Value
    if(-not $script:visibleRecordByLine.ContainsKey($line)) {return}
    $rec=$script:visibleRecordByLine[$line]
    $map=@{1=0;2=2;6=6;7=7;8=1}
    if(-not $map.ContainsKey($idx)) {return}
    if($idx -eq 8 -and -not $advanced.Checked) {return}
    $new=if($null -eq $row.Cells[$idx].Value){''}else{[string]$row.Cells[$idx].Value}
    if($idx -eq 6 -or $idx -eq 7) {$new=Normalize-Address $new}
    $old=[string]$rec.Values[$map[$idx]]
    if($new -ceq $old) {return}
    Snapshot
    if($idx -eq 1) {
        # Moving a device to another room also moves every service in that accessory.
        foreach($item in $tree.SelectedNode.Tag) {$item.Values[0]=$new}
    } elseif($idx -eq 2 -or $idx -eq 8) {
        $key=Service-Key $rec.Values
        foreach($item in $script:project) {
            if((Service-Key $item.Values) -ceq $key) {$item.Values[$map[$idx]]=$new}
        }
    } else {$rec.Values[$map[$idx]]=$new}
    Queue-Rebuild $rec.Line
})
function Do-Undo([bool]$redo) {
    $from=if($redo){$script:redo}else{$script:undo}
    $to=if($redo){$script:undo}else{$script:redo}
    if($from.Count -eq 0) {return}
    $to.Push((Clone-Records $script:project))
    $script:project=$from.Pop()
    Rebuild-Tree 0
}
$tree.Add_AfterSelect({if(-not $script:drawing){Queue-DetailRefresh}})
$filter.Add_TextChanged({if(-not $script:drawing){Rebuild-Tree 0}})
$advanced.Add_CheckedChanged({$grid.Columns[8].Visible=$advanced.Checked})
$undoBtn.Add_Click({Do-Undo $false})
$redoBtn.Add_Click({Do-Undo $true})
$form.Add_KeyDown({
    if($_.Control -and $_.KeyCode -eq 'Z') {Do-Undo $false;$_.SuppressKeyPress=$true}
    if($_.Control -and $_.KeyCode -eq 'Y') {Do-Undo $true;$_.SuppressKeyPress=$true}
})
$restoreBtn.Add_Click({
    if($null -eq $tree.SelectedNode -or $tree.SelectedNode.Level -ne 1) {return}
    if([Windows.Forms.MessageBox]::Show('将此设备恢复到刚导入时的配置？','恢复','YesNo') -ne 'Yes') {return}
    Snapshot
    $line=$tree.SelectedNode.Tag[0].Line
    foreach($rec in $tree.SelectedNode.Tag) {
        foreach($base in $script:baseline) {
            if($base.Line -eq $rec.Line) {
                $rec.Values=[string[]]@($base.Values)
                $rec.Group=$base.Group
                break
            }
        }
    }
    Rebuild-Tree $line
})
$loadCsv.Add_Click({Choose-Input 'csv'})
$loadXlsx.Add_Click({Choose-Input 'xlsx'})
$saveXlsx.Add_Click({Do-Export 'xlsx'})
$saveCsv.Add_Click({Do-Export 'csv'})
$previewBtn.Add_Click({$tabs.SelectedTab=$pagePreview})
$reportBtn.Add_Click({
    if(-not $script:projectPath) {return}
    $c=Show-Issues
    $folder=Join-Path $script:outputBase '检查报告'
    [void][IO.Directory]::CreateDirectory($folder)
    $target=Unique-Target $folder ((New-ExportName $script:projectPath)+'_检查报告') '.txt'
    $report=@("来源：$script:projectPath","错误：$($c.Errors.Count)","提醒：$($c.Warnings.Count)",'')
    $report+=@($c.Errors);$report+=@($c.Warnings)
    [IO.File]::WriteAllLines($target,[string[]]$report,[Text.UTF8Encoding]::new($true))
    Write-Log ('检查报告：'+$target)
    [void][Windows.Forms.MessageBox]::Show("报告已生成："+[Environment]::NewLine+$target,'全面检测')
})
$issueList.Add_DoubleClick({
    if([string]$issueList.SelectedItem -match '第\s+(\d+)\s+行') {
        $node=Find-Line ([int]$Matches[1])
        if($null -ne $node) {$tree.SelectedNode=$node}
    }
})
function Pick-Choice($choices,[string]$title) {
    $dialog=New-Object Windows.Forms.Form
    $dialog.Text=$title
    $dialog.Size=New-Object Drawing.Size(530,380)
    $dialog.StartPosition='CenterParent'
    $list=New-Object Windows.Forms.ListBox
    $list.Location=New-Object Drawing.Point(12,12)
    $list.Size=New-Object Drawing.Size(485,270)
    foreach($choice in $choices) {[void]$list.Items.Add($choice.Label)}
    $dialog.Controls.Add($list)
    $ok=New-Object Windows.Forms.Button
    $ok.Text='确定'
    $ok.Location=New-Object Drawing.Point(396,298)
    $ok.DialogResult=[Windows.Forms.DialogResult]::OK
    $dialog.Controls.Add($ok)
    $dialog.AcceptButton=$ok
    try {
        if($dialog.ShowDialog($form) -eq [Windows.Forms.DialogResult]::OK -and $list.SelectedIndex -ge 0) {
            return $choices[$list.SelectedIndex]
        }
    } finally {$dialog.Dispose()}
    return $null
}
$groupBtn.Add_Click({
    $selected=$tree.SelectedNode
    if($null -eq $selected -or $selected.Level -ne 1) {return}
    $choices=New-Object 'System.Collections.Generic.List[object]'
    foreach($sibling in $selected.Parent.Nodes) {
        if([object]::ReferenceEquals($sibling,$selected)) {continue}
        [void]$choices.Add([pscustomobject]@{Label=$sibling.Text;Rows=$sibling.Tag})
    }
    if($choices.Count -eq 0) {[void][Windows.Forms.MessageBox]::Show('同一房间内没有其他可组合的配件。');return}
    $picked=Pick-Choice $choices '选择要加入当前配件的服务'
    if($null -eq $picked) {return}
    $left=@($selected.Tag | ForEach-Object {$_})
    $right=@($picked.Rows | ForEach-Object {$_})
    $all=@($left)+@($right)
    $ids=Unique-Strings @($all | ForEach-Object {$_.Values[1]})
    if($ids.Length -gt 1) {
        [void][Windows.Forms.MessageBox]::Show('这些服务含有不同的非空 UUID，请先核对；不能自动合并。','UUID 冲突')
        return
    }
    $label=''
    foreach($rec in $all) {if($rec.Group -ne '') {$label=$rec.Group;break}}
    if(-not $label) {$label='组合'+[Guid]::NewGuid().ToString('N').Substring(0,8)}
    Snapshot
    $focus=$left[0].Line
    foreach($rec in $all) {$rec.Group=$label}
    Rebuild-Tree $focus
})
$separateBtn.Add_Click({
    $selected=$tree.SelectedNode
    if($null -eq $selected -or $selected.Level -ne 1) {return}
    $services=New-Object 'System.Collections.Generic.List[object]'
    $found=@{}
    foreach($rec in $selected.Tag) {
        $key=Service-Key $rec.Values
        if($found.ContainsKey($key)) {continue}
        $found[$key]=$true
        [void]$services.Add([pscustomobject]@{
            Label=$rec.Values[2]+' ('+$rec.Values[3]+')'
            Key=$key
            Line=$rec.Line
        })
    }
    if($services.Count -le 1) {[void][Windows.Forms.MessageBox]::Show('当前配件只有一个服务，无需解除组合。');return}
    $picked=Pick-Choice $services '选择要解除组合的服务'
    if($null -eq $picked) {return}
    if([Windows.Forms.MessageBox]::Show('解除组合会清空被拆分服务的 UUID，可能改变 HomeKit 中的配件身份。继续？','确认解除组合','YesNo','Warning') -ne 'Yes') {return}
    Snapshot
    foreach($rec in $script:project) {
        if((Service-Key $rec.Values) -ceq $picked.Key) {
            $rec.Group=''
            $rec.Values[1]=''
        }
    }
    Rebuild-Tree $picked.Line
})
foreach($element in @($pageFiles,$pagePreview)) {
    $element.AllowDrop=$true
    $element.Add_DragEnter({
        if($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) {$_.Effect='Copy'}
    })
    $element.Add_DragDrop({
        $paths=@($_.Data.GetData([Windows.Forms.DataFormats]::FileDrop))
        if($paths.Count -gt 0) {
            try {Load-File $paths[0]}
            catch {[void][Windows.Forms.MessageBox]::Show($_.Exception.Message,'拖入文件失败')}
        }
    })
}
$form.Add_FormClosing({
    if(Has-Unsaved) {
        if([Windows.Forms.MessageBox]::Show('当前修改尚未导出，是否仍要关闭？','关闭确认','YesNo','Warning') -ne 'Yes') {$_.Cancel=$true}
    }
})
if($GuiSmokeTest) {
    try {
        $sample=Join-Path $PSScriptRoot '测试数据\三行示例.csv'
        Load-File $sample
        if($tree.Nodes.Count -ne 1) {throw '预览设备树房间数量异常。'}
        if($tree.Nodes[0].Nodes.Count -ne 2) {throw '预览设备树服务数量异常。'}
        if($grid.Rows.Count -lt 1) {throw '右侧设备详情没有加载。'}
        $form.Show()
        $form.Activate()
        [Windows.Forms.Application]::DoEvents()
        $tabs.SelectedTab=$pagePreview
        [Windows.Forms.Application]::DoEvents()
        $humidifierSample=Join-Path $PSScriptRoot '测试数据\组合加湿机_合成样本.csv'
        Load-File $humidifierSample
        if($tree.Nodes.Count -ne 3) {throw '双服务加湿机：房间数量异常。'}
        Write-Host ("界面尺寸诊断：form={0}; tabs={1}; preview={2}; split={3}; panel2={4}; grid={5}" -f $form.ClientSize,$tabs.Size,$pagePreview.Size,$split.Size,$split.Panel2.Size,$grid.Size)
        # Assert actual non-overlap in on-screen coordinates; mere RowCount=2
        # would pass even when the first row and column headers are covered.
        if(-not [object]::ReferenceEquals($grid.Parent,$rightLayout)) {throw '参数表格未放置于独立布局行。'}
        if(-not [object]::ReferenceEquals($heading.Parent,$rightLayout)) {throw '设备标题未放置于独立布局行。'}
        if(-not [object]::ReferenceEquals($buttons.Parent,$rightLayout)) {throw '操作按钮未放置于独立布局行。'}
        $actionBottom=$buttons.PointToScreen([Drawing.Point]::Empty).Y + $buttons.Height
        $gridTop=$grid.PointToScreen([Drawing.Point]::Empty).Y
        if($gridTop -lt $actionBottom - 1) {
            throw "表格被工具栏遮挡：gridTop=$gridTop, actionsBottom=$actionBottom"
        }
        if(-not $grid.ColumnHeadersVisible -or $grid.ColumnHeadersHeight -lt 35) {
            throw '参数标题行没有正确显示。'
        }
        if(-not $grid.AllowUserToResizeColumns -or $grid.Columns[5].Resizable -ne [Windows.Forms.DataGridViewTriState]::True) {
            throw '列宽不能通过拖动表头分隔线调整。'
        }
        Write-Host ("布局覆盖检测通过：操作栏底={0}, 表格顶={1}, 标题行高度={2}" -f $actionBottom,$gridTop,$grid.ColumnHeadersHeight)
        foreach($testRoom in $tree.Nodes) {
            if($testRoom.Nodes.Count -ne 1) {throw '双服务加湿机：配件数量异常。'}
            $humidifierNode=$testRoom.Nodes[0]
            if($humidifierNode.Tag.Count -ne 2) {throw "双服务加湿机：配件 $($humidifierNode.Text) 应有 2 行。"}
            $tree.SelectedNode=$humidifierNode
            Show-Device
            if($grid.Rows.Count -ne 2) {
                throw "双服务加湿机：$($humidifierNode.Text) 的 DataGridView 实际 $($grid.Rows.Count) 行，而 CSV 有 2 行。"
            }
            [Windows.Forms.Application]::DoEvents()
            if(-not $grid.Visible -or $grid.Height -lt 100 -or $split.Panel2.Width -lt 480) {
                throw "双服务加湿机：详情表格不可见或尺寸异常。可见=$($grid.Visible)；尺寸=$($grid.Size)"
            }
            $firstRowRect=$grid.GetRowDisplayRectangle(0,$false)
            if($firstRowRect.Y -lt $grid.ColumnHeadersHeight - 1) {
                throw "第一行被表头覆盖：首行起点 $($firstRowRect.Y), 表头 $($grid.ColumnHeadersHeight)"
            }
            if($grid.DisplayedRowCount($false) -lt 2 -or $grid.GetRowDisplayRectangle(0,$false).Height -lt 15) {
                throw "双服务加湿机：$($humidifierNode.Text) 的表格行没有实际显示。表格=$($grid.Bounds)；显示行数=$($grid.DisplayedRowCount($false))"
            }
            if(-not $grid.ColumnHeadersVisible -or $grid.Columns.Count -lt 9) {
                throw '双服务加湿机：详情表头没有显示。'
            }
        }
        # Regression: switch between the three humidifiers repeatedly and edit an
        # address. DataGridView must not throw SetCurrentCellAddressCore on Rows.Clear,
        # and the first two service rows must remain visible afterward.
        for($cycle=0;$cycle -lt 12;$cycle++) {
            foreach($roomToSelect in $tree.Nodes) {
                $tree.SelectedNode=$roomToSelect.Nodes[0]
                [Windows.Forms.Application]::DoEvents()
                if($grid.Rows.Count -ne 2) {throw '切换加湿机时设备详情变空白。'}
                if([string]$grid.Rows[0].Cells[3].Value -cne 'Switch' -or
                   [string]$grid.Rows[1].Cells[3].Value -cne 'HumiditySensor') {
                    throw '加湿机的 Switch/HumiditySensor 服务没有完整显示。'
                }
            }
        }
        $tree.SelectedNode=$tree.Nodes[0].Nodes[0]
        [Windows.Forms.Application]::DoEvents()
        $targetCell=$grid.Rows[0].Cells[6]
        $grid.CurrentCell=$targetCell
        if(-not $grid.BeginEdit($false)) {throw '无法开始编辑加湿机组地址。'}
        $grid.EditingControl.Text='5 / 1 / 99'
        if(-not $grid.EndEdit()) {throw '无法完成加湿机组地址编辑。'}
        [Windows.Forms.Application]::DoEvents()
        [Windows.Forms.Application]::DoEvents()
        if($grid.Rows.Count -ne 2 -or [string]$grid.Rows[0].Cells[6].Value -cne '5/1/99') {
            throw '地址编辑后加湿机详情未恢复，或新组地址没有正确保存。'
        }
        if($grid.Columns[3].Width -lt 180 -or $grid.Columns[5].Width -lt 280 -or
           $grid.DefaultCellStyle.WrapMode -ne [Windows.Forms.DataGridViewTriState]::True) {
            throw '列宽或文本换行设置没有生效。'
        }
        if((Normalize-Address '1 2 3') -cne '1/2/3') {throw '组地址空格输入转换失败。'}
        Write-Host 'GUI 自检通过：导入样本、房间树、设备详情和组地址规范化。'
        exit 0
    } catch {
        Write-Host ('GUI 自检失败：'+$_.Exception.Message)
        exit 1
    } finally {$form.Dispose()}
}
[void]$form.ShowDialog()
