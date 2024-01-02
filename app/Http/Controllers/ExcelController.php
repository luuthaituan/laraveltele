<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use Illuminate\Support\Facades\DB;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use GuzzleHttp\Client;

class ExcelController extends Controller
{
    public function index(Request $request)
    {
        try {
            // Kiểm tra nếu là request POST
            if ($request->isMethod('post')) {
                // Thực hiện truy vấn
                $result = DB::select("SELECT * FROM local.test");

                // Gửi kết quả về trình duyệt
                if (count($result) > 0) {
                    // Xuất ra file Excel và gửi chi tiết kết quả về Telegram
                    $excelFilePath = $this->generateExcelFile($result, 'query_result.xlsx');
                    $this->sendTelegramDocument($excelFilePath);

                    return view('index', ['result' => $result, 'excelFilePath' => $excelFilePath]);
                } else {
                    // Gửi thông báo không có dữ liệu về Telegram
                    $this->sendTelegramMessage('No data found in the database.');

                    return view('index', ['result' => null]);
                }
            } else {
                // Logic khi là request GET
                return view('index', ['result' => null]);
            }
        } catch (\Exception $e) {
            // Xử lý lỗi và gửi về trình duyệt
            return view('index', ['result' => null, 'error' => $e->getMessage()]);
        }
    }

    private function sendTelegramMessage($message)
    {
        // Thay thế YOUR_BOT_TOKEN và YOUR_CHAT_ID bằng thông tin của bot và chat của bạn
        $botToken = '6531445104:AAFviAc-HUb2LQ0hL0qX_mEu3NvSonO2D8k';
        $chatId = '2065288078';

        // Tạo URL để gửi tin nhắn
        $url = "https://api.telegram.org/bot{$botToken}/sendMessage";

        // Dữ liệu để gửi
        $data = [
            'chat_id' => $chatId,
            'text' => $message,
        ];

        // Sử dụng GuzzleHTTP để gửi HTTP POST request
        $client = new Client();
        $client->post($url, ['json' => $data]);
    }

    private function sendTelegramDocument($filePath)
    {
        // Thay thế YOUR_BOT_TOKEN và YOUR_CHAT_ID bằng thông tin của bot và chat của bạn
        $botToken = '6531445104:AAFviAc-HUb2LQ0hL0qX_mEu3NvSonO2D8k';
        $chatId = '2065288078';

        // Tạo URL để gửi tài liệu
        $url = "https://api.telegram.org/bot{$botToken}/sendDocument";

        // Sử dụng GuzzleHTTP để gửi HTTP POST request
        $client = new Client();
        $client->request('POST', $url, [
            'multipart' => [
                [
                    'name' => 'chat_id',
                    'contents' => $chatId,
                ],
                [
                    'name' => 'document',
                    'contents' => fopen($filePath, 'r'),
                    'filename' => basename($filePath),
                ],
            ],
        ]);
    }

    private function generateExcelFile($data, $fileName)
    {
        // Tạo đối tượng Spreadsheet
        $spreadsheet = new Spreadsheet();

        // Tạo sheet mới
        $sheet = $spreadsheet->getActiveSheet();

        // Ghi dữ liệu từ mảng vào sheet
        $rowIndex = 1;
        foreach ($data as $row) {
            $colIndex = 1;
            foreach ($row as $value) {
                $sheet->setCellValueByColumnAndRow($colIndex, $rowIndex, $value);
                $colIndex++;
            }
            $rowIndex++;
        }

        // Tạo writer để ghi vào file Excel
        $writer = new Xlsx($spreadsheet);

        // Lưu file Excel
        $excelFilePath = storage_path($fileName);
        $writer->save($excelFilePath);

        return $excelFilePath;
    }
}
