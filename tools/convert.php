<?php
require __DIR__.'/vendor/autoload.php';
use Carbon\Carbon;
use Tightenco\Collect\Support\Collection;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as Xlsx;

# メモリ不足エラーが発生する場合は設定する
# ini_set("memory_limit", "300M");

# PCR検査数＋陽性者数
$summaries = setInspectionJson();
file_put_contents(__DIR__.'/../data/inspections.json', json_encode($summaries, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_NUMERIC_CHECK));

# 患者状況(日付・市町村・年代)
$patients = setPatientJson();
file_put_contents(__DIR__.'/../data/patients.json', json_encode($patients, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE | JSON_NUMERIC_CHECK));

/*
 * PCR検査人数＋陽性者人数の配列を作成
 *
 * @return [    "date"  =>  インポートした日付
 *              "row"   =>  [
 *                  "date"                  => 日付
 *                  "week_count"            => 週数
 *                  "inspection_count"      => 検査件数
 *                  "positive_person_count" => 陽性者数
 *              ]
 *          ]
 */
function setInspectionJson(): array
{
    $data = xlsxToArray(__DIR__.'/downloads/m-covid-kensa.xlsx', 'PCR検査(件数詳細) ', 'A10:E371', 'A9:E9');
    return [
        'date' => Carbon::today()->format('Y-m-d'),
        'data' => $data->filter(function ($row){
            $today = Carbon::today();

            # 行の日付整形
            $row_date =  explode("日", Carbon::today()->year."-".str_replace('月', '-', $row['日付']))[0];
            $date =  createDate($row_date);

            # 現在日付より過去のみ対象
            return $date !== null && $date->lt($today);
        })->map(function ($row) {
            $row_date =  explode("日", Carbon::today()->year."-".str_replace('月', '-', $row['日付']))[0];
            $date =  createDate($row_date);

            return [
                "date" => $date !== null ? $date->format('Y-m-d') : "",
                "week_count" => $row["週数"],
                "inspection_count" => str_replace([',', ' '], '', $row["検査件数"]),
                "positive_person_count" => str_replace([',', ' '], '', $row["陽性数"])
            ];
        })
    ];
}

/*
 * 新規患者情報の配列を作成
 *
 * @return [    "date"  =>  インポートした日付
 *              "row"   =>  [
 *                  "release_date"  => 公表日
 *                  "age"           => 患者の年代
 *                  "gender"        => 患者の性別
 *                  "business_type" => 患者の職業
 *                  "city"          => 市町村
 *                  "onset_date"    => 発症日
 *                  "positive_date" => 陽性判明日
 *                  "state"         => 療養状況
 *              ]
 *          ]
 */
function setPatientJson(): array
{
    $data = xlsxToArray(__DIR__.'/downloads/m-covid-kanja.xlsx', '患者状況一覧（HP掲載）', 'A4:I20000', 'A3:I3');
    return [
        'date' => Carbon::today()->format('Y-m-d'),
        'data' => $data->filter(function ($row){
            $today = Carbon::today();

            # 行の日付整形
            $release_date =  createDate($row['公表_年月日']);

            # 現在日付より過去のみ対象
            return $release_date !== null && $release_date->lt($today);
        })->map(function ($row) {
            $release_date =  createDate($row['公表_年月日']);
            $onset_date =  createDate($row['患者_発症日']);
            $positive_date =  createDate($row['陽性判明_年月日']);

            return [
                "release_date" => $release_date !== null ? $release_date->format('Y-m-d') : "",
                "age" => $row["患者_年代"],
                "gender" => $row["患者_性別"],
                "business_type" => $row["患者_職業"],
                "city" => $row["患者_居住地"],
                "onset_date" => $onset_date !== null ? $onset_date->format('Y-m-d') : "",
                "positive_date" => $positive_date !== null ? $positive_date->format('Y-m-d') : "",
                "state" => $row["患者_療養状況"]
            ];
        })
    ];
}

/*
 * 日付の整形
 *
 * @param string $date 加工前の日付
 * @return ?Carbon 加工後の日付
 */
function createDate(?string $date): ?Carbon
{
    if ($date === null) {
        return  null;
    }

    # 日付の整形
    $explode_date =  explode("/", $date);
    if (count($explode_date) === 3){
        return Carbon::create($explode_date[0], $explode_date[1], $explode_date[2]);
    }

    return null;
}

/*
 * Excelファイルを一行ずつ読み込みCollectionに格納
 *
 * @param string $path ファイル名
 * @param string $sheet_name シート名
 * @param string $range データ部の範囲
 * @param string|null $header_range ヘッダー部の範囲
 * @return Collection $data ExcelデータをCollectionで整形した結果
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