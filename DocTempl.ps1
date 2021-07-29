function ProcessDocFile
{
    param($dir, $name, $fields, $word)
    $doc = $word.Documents.Open($dir + "\" + $name);
    $vars = $doc.Variables
    $path = (Get-Location).Path;

    $resName = $name;
    for ($i = 0; $i -lt $fields.Count; $i++) {
        $field = $fields[$i];
        $resName = $resName.Replace("$" + $field.Name + "$", $field.Value);

        $vars[$field.Name].Value = $field.Value;
    }

    $doc.Fields.Update();
    $doc.SaveAs($path + '\' + $resName);
    $doc.Close(0);
}

Write-Host "Формирование документов из шаблонов", $args[0];

$mode = 0
$files = @();
$fields = @();
$cnt = $args[0];
if ($cnt -eq "" ) {
  $cnt = "DocTempl.txt".
}

$ini = Get-Content -Encoding utf8 $cnt;
for ($i = 0; $i -lt $ini.Count; $i++) {
    $line = $ini[$i].trim();
    $line = $ini[$i].trim();
    if ($line -eq "") {
    } elseif ($line.ToLower() -eq "[шаблоны]") {
        $mode = 1;
    } elseif ($line.ToLower() -eq "[поля]") {
        $mode = 2;
    } else {
        if ($mode -eq 1) {
            $files += $line;  
        } elseif ($mode -eq 2) {
            $fld = $line.Split("=", 2);
            $fields += [PSCustomObject]@{ 
                Name = $fld[0] 
                Value = $fld[1] 
            };
        }
    }
}

$word = New-Object -ComObject Word.Application
$word.DisplayAlerts = 0

for ($i = 0; $i -lt $files.Length; $i++) {
    $file = $files[$i];
    if (Test-Path $file) {
        $item = (Get-Item $file);
        $p = ProcessDocFile -dir $item.DirectoryName -name $item.Name -fields $fields -word $word;
    } else {
        Write-Host "File not exists", $file;
    }
}

$word.Quit();
