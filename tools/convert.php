<?php
require __DIR__.'/vendor/autoload.php';
use Carbon\Carbon;
use Tightenco\Collect\Support\Collection;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Xlsx;

# PCR検査数＋陽性者数
$summaries = setSummaryJson();

# 患者状況(日付・市町村・年代)
//$patients = setPatientJson();

file_put_contents(__DIR__.'/../data/summaries.json', json_encode($summaries, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_NUMERIC_CHECK));
//file_put_contents(__DIR__.'/../data/patients.json', json_encode($patients, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_NUMERIC_CHECK));

function setSummaryJson(): array
{
    $data = xlsxToArray(__DIR__.'/downloads/m-covid-kensa.xlsx', 'PCR検査(件数詳細) ', 'A10:E371', 'A9:E9');
    return [
        'date' => Carbon::today()->format('Y-m-d'),
        'data' => $data->filter(function ($row){
            $today = Carbon::today();

            # 行の日付整形
            $date =  explode("日", Carbon::today()->year."-".str_replace('月', '-', $row['日付']));
            $arrayDate = explode("-", $date[0]);
            $rowDate = Carbon::create($arrayDate[0], $arrayDate[1], $arrayDate[2]);

            # 現在日付より過去のみ対象
            return $rowDate->lt($today);
        })->map(function ($row) {
            $date =  explode("日", Carbon::today()->year."-".str_replace('月', '-', $row['日付']));
            $arrayDate = explode("-", $date[0]);

            return [
                "date" => Carbon::create($arrayDate[0], $arrayDate[1], $arrayDate[2])->format('Y-m-d'),
                "week_count" => $row["週数"],
                "inspection_count" => str_replace([',', ' '], '', $row["検査件数"]),
                "positive_person_count" => str_replace([',', ' '], '', $row["陽性数"])
            ];
        })
    ];
}

function setPatientJson() {

}

/*
 * Excelファイルを一行ずつ読み込みCollectionに格納
 *
 * @param string $path
 * @param string $sheet_name
 * @param string $range
 * @param string|null $header_range
 * @return Collection $data
 */
function xlsxToArray(string $path, string $sheet_name, string $range, $header_range = null): Collection
{
    $reader = new Xlsx();
    $spreadsheet = $reader->load($path);
    $sheet = $spreadsheet->getSheetByName($sheet_name);
    $data =  new Collection($sheet->rangeToArray($range));
    $data = $data->filter(function ($row){
        return $row !== "";
    })->map(function ($row) {
        return new Collection($row);
    });
    if ($header_range !== null) {
        $headers = xlsxToArray($path, $sheet_name, $header_range)[0];
        return $data->map(function ($row) use($headers){
            return $row->mapWithKeys(function ($cell, $idx) use($headers){
                return [
                    $headers[$idx] => $cell
                ];
            });
        });
    }

    return $data;
}