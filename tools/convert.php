<?php
include('./vendor/autoload.php');
use phpoffi\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use Carbon\Carbon;
use Tightenco\Collect\Support\Collection;

# PCR検査数＋陽性者数
$summaries = setSummaryJson();

# 患者状況(日付・市町村・年代)
$patients = setPatientJson();

file_put_contents(__DIR__.'/../data/summaries.json', json_encode($summaries, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_NUMERIC_CHECK));
file_put_contents(__DIR__.'/../data/patients.json', json_encode($patients, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_NUMERIC_CHECK));

function setSummaryJson() {
    $data = xlsxToArray(__DIR__.'/downloads/コールセンター相談件数-RAW.xlsx', 'Sheet1', 'A2:E100', 'A1:E1');
    return [
        'date' => xlsxToArray(__DIR__.'/downloads/コールセンター相談件数-RAW.xlsx', 'Sheet1', 'H1')[0][0],
        'data' => $data->filter(function ($row) {
            return $row['曜日'] && $row['17-21時'];
        })->map(function ($row) {
            $date = '2020-'.str_replace(['月', '日'], ['-', ''], $row['日付']);
            $carbon = Carbon::parse($date);
            $row['日付'] = $carbon->format('Y-m-d').'T08:00:00.000Z';
            $row['date'] = $carbon->format('Y-m-d');
            $row['w'] = $carbon->format('w');
            $row['short_date'] = $carbon->format('m/d');
            $row['小計'] = array_sum([
                $row['9-13時'] ?? 0,
                $row['13-17時'] ?? 0,
                $row['17-21時'] ?? 0,
            ]);
            return $row;
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
    $reader = new XlsxReader();
    $spreadsheet = $reader->load($path);
    $sheet = $spreadsheet->getSheetByName($sheet_name);
    $data =  new Collection($sheet->rangeToArray($range));
    $data = $data->map(function ($row) {
        return new Collection($row);
    });
    if ($header_range !== null) {
        $headers = xlsxToArray($path, $sheet_name, $header_range)[0];
        // TODO check same columns length
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