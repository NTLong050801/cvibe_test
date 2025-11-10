<?php

namespace App\Jobs;

use App\Models\VehicleRegistration;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class ExportVehicleRegistrationsJob implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    protected $thang;
    protected $nam;

    /**
     * Số lần thử lại nếu job thất bại
     */
    public $tries = 3;

    /**
     * Timeout cho job (giây)
     */
    public $timeout = 600;

    /**
     * Create a new job instance.
     */
    public function __construct($thang, $nam)
    {
        $this->thang = $thang;
        $this->nam = $nam;
    }

    /**
     * Execute the job.
     */
    public function handle(): string
    {
        // Load template Excel
        $templatePath = storage_path('Vexe_Template.xlsx');

        if (!file_exists($templatePath)) {
            throw new \Exception('Template file not found: ' . $templatePath);
        }

        $spreadsheet = IOFactory::load($templatePath);
        $sheet = $spreadsheet->getActiveSheet();

        // Lấy dữ liệu từ database
        $registrations = VehicleRegistration::where('thang', $this->thang)
            ->where('nam', $this->nam)
            ->orderBy('stt')
            ->get();

        if ($registrations->isEmpty()) {
            throw new \Exception('Không có dữ liệu để export');
        }

        // Tính số nhóm cần thiết (mỗi nhóm có 3 vé)
        $totalTickets = $registrations->count();
        $groupsNeeded = ceil($totalTickets / 3);

        // TÌM VỊ TRÍ THỰC TẾ của các nhóm trong template
        $groupPositions = [];
        $highestRow = $sheet->getHighestRow();

        for ($row = 1; $row <= $highestRow; $row++) {
            $value = $sheet->getCell('A' . $row)->getValue();
            if ($value === 'VÉ GỬI XE') {
                $groupPositions[] = $row;
            }
        }

        $templateGroupCount = count($groupPositions);

        // Số nhóm tối đa trên 1 sheet = số nhóm có sẵn trong template
        $maxGroupsPerSheet = $templateGroupCount;

        // Đặt tên cho sheet đầu tiên
        $sheet->setTitle('Trang 1');

        // Nếu cần nhiều hơn template, tạo thêm sheet
        $sheetsNeeded = ceil($groupsNeeded / $maxGroupsPerSheet);
        $allSheets = [$sheet];

        if ($sheetsNeeded > 1) {
            // Clone sheet gốc để tạo thêm sheets
            for ($sheetIndex = 1; $sheetIndex < $sheetsNeeded; $sheetIndex++) {
                $newSheet = clone $sheet;
                $newSheet->setTitle('Trang ' . ($sheetIndex + 1));
                $spreadsheet->addSheet($newSheet);
                $allSheets[] = $newSheet;
            }
        }

        // Xử lý từng sheet
        foreach ($allSheets as $sheetIndex => $currentSheet) {
            // Tính số nhóm cần thiết cho sheet này
            $groupsInThisSheet = min($maxGroupsPerSheet, $groupsNeeded - ($sheetIndex * $maxGroupsPerSheet));

            if ($groupsInThisSheet <= 0) {
                break;
            }

            // Tìm lại group positions cho sheet hiện tại
            $sheetGroupPositions = [];
            $sheetHighestRow = $currentSheet->getHighestRow();

            for ($row = 1; $row <= $sheetHighestRow; $row++) {
                $value = $currentSheet->getCell('A' . $row)->getValue();
                if ($value === 'VÉ GỬI XE') {
                    $sheetGroupPositions[] = $row;
                }
            }

            $sheetTemplateGroupCount = count($sheetGroupPositions);

            // Nếu template không đủ nhóm cho sheet này, DUPLICATE thêm
            if ($sheetTemplateGroupCount < $groupsInThisSheet) {
                $sourceGroupIndex = ($sheetTemplateGroupCount > 1) ? 1 : 0;
                $sourceGroupStart = $sheetGroupPositions[$sourceGroupIndex];

                $groupHeight = ($sheetTemplateGroupCount > 1)
                    ? ($sheetGroupPositions[1] - $sheetGroupPositions[0])
                    : 9;

                $lastGroupEnd = $sheetGroupPositions[$sheetTemplateGroupCount - 1] + $groupHeight - 1;

                // Lưu thông tin ảnh từ nhóm mẫu
                $sourceImages = [];
                foreach ($currentSheet->getDrawingCollection() as $drawing) {
                    $coords = $drawing->getCoordinates();
                    preg_match('/([A-Z]+)(\d+)/', $coords, $matches);

                    if ($matches) {
                        $row = (int)$matches[2];
                        if ($row >= $sourceGroupStart && $row < ($sourceGroupStart + $groupHeight)) {
                            $sourceImages[] = [
                                'drawing' => $drawing,
                                'offset' => $row - $sourceGroupStart,
                            ];
                        }
                    }
                }

                // Lưu merged cells vào array TRƯỚC KHI iterate
                $mergedCellsArray = $currentSheet->getMergeCells();

                // Copy merged cells TRƯỚC để giữ nguyên structure
                $mergedCellsToAdd = [];
                foreach ($mergedCellsArray as $mergeCell) {
                    if (preg_match('/([A-Z]+)(\d+):([A-Z]+)(\d+)/', $mergeCell, $matches)) {
                        $startCol = $matches[1];
                        $startRow = (int)$matches[2];
                        $endCol = $matches[3];
                        $endRow = (int)$matches[4];

                        if ($startRow >= $sourceGroupStart && $startRow < $sourceGroupStart + $groupHeight) {
                            $rowOffset = $startRow - $sourceGroupStart;
                            $rowEndOffset = $endRow - $sourceGroupStart;

                            for ($groupIndex = $sheetTemplateGroupCount; $groupIndex < $groupsInThisSheet; $groupIndex++) {
                                $newGroupStart = $lastGroupEnd + 1 + (($groupIndex - $sheetTemplateGroupCount) * $groupHeight);
                                $newStartRow = $newGroupStart + $rowOffset;
                                $newEndRow = $newGroupStart + $rowEndOffset;
                                $mergedCellsToAdd[] = $startCol . $newStartRow . ':' . $endCol . $newEndRow;
                            }
                        }
                    }
                }

                // Duplicate thêm các nhóm còn thiếu
                for ($groupIndex = $sheetTemplateGroupCount; $groupIndex < $groupsInThisSheet; $groupIndex++) {
                    $newGroupStart = $lastGroupEnd + 1 + (($groupIndex - $sheetTemplateGroupCount) * $groupHeight);

                    // Copy từng dòng từ nhóm mẫu
                    for ($rowOffset = 0; $rowOffset < $groupHeight; $rowOffset++) {
                        $sourceRow = $sourceGroupStart + $rowOffset;
                        $targetRow = $newGroupStart + $rowOffset;

                        // Copy row height
                        $currentSheet->getRowDimension($targetRow)
                            ->setRowHeight($currentSheet->getRowDimension($sourceRow)->getRowHeight());

                        // Copy từng cell với style đầy đủ
                        foreach (range('A', 'I') as $col) {
                            $sourceCellCoord = $col . $sourceRow;
                            $targetCellCoord = $col . $targetRow;

                            // Copy value (chỉ label, không copy data B, E, H)
                            if (!in_array($col, ['B', 'E', 'H'])) {
                                $currentSheet->setCellValue($targetCellCoord, $currentSheet->getCell($sourceCellCoord)->getValue());
                            }

                            // Copy style bằng duplicateStyle
                            $currentSheet->duplicateStyle(
                                $currentSheet->getStyle($sourceCellCoord),
                                $targetCellCoord
                            );
                        }
                    }

                    // Copy column width (chỉ cần copy 1 lần)
                    if ($groupIndex == $sheetTemplateGroupCount) {
                        foreach (range('A', 'I') as $col) {
                            $currentSheet->getColumnDimension($col)
                                ->setWidth($currentSheet->getColumnDimension($col)->getWidth());
                        }
                    }

                    // Clone ảnh cho nhóm mới
                    foreach ($sourceImages as $imageInfo) {
                        $sourceDrawing = $imageInfo['drawing'];
                        $offset = $imageInfo['offset'];

                        $newDrawing = clone $sourceDrawing;
                        $newRow = $newGroupStart + $offset;

                        preg_match('/([A-Z]+)(\d+)/', $sourceDrawing->getCoordinates(), $matches);
                        $col = $matches[1];

                        $newDrawing->setCoordinates($col . $newRow);
                        $newDrawing->setWorksheet($currentSheet);
                    }
                }

                // Merge cells cho các nhóm mới
                foreach ($mergedCellsToAdd as $mergeRange) {
                    try {
                        $currentSheet->mergeCells($mergeRange);
                    } catch (\Exception $e) {
                        // Ignore nếu đã merge hoặc conflict
                    }
                }
            }
        }

        // ĐIỀN DỮ LIỆU vào các vé trên tất cả sheets
        $ticketIndex = 0;
        foreach ($registrations as $registration) {
            // Tính sheet nào và group nào
            $sheetIndex = floor($ticketIndex / ($maxGroupsPerSheet * 3));
            $ticketInSheet = $ticketIndex % ($maxGroupsPerSheet * 3);
            $groupInSheet = floor($ticketInSheet / 3);
            $positionInGroup = $ticketInSheet % 3;

            if ($sheetIndex >= count($allSheets)) {
                break;
            }

            $currentSheet = $allSheets[$sheetIndex];

            // Tìm lại group positions cho sheet hiện tại
            $sheetGroupPositions = [];
            for ($row = 1; $row <= $currentSheet->getHighestRow(); $row++) {
                $value = $currentSheet->getCell('A' . $row)->getValue();
                if ($value === 'VÉ GỬI XE') {
                    $sheetGroupPositions[] = $row;
                }
            }

            if ($groupInSheet >= count($sheetGroupPositions)) {
                break;
            }

            $groupStartRow = $sheetGroupPositions[$groupInSheet];
            $colOffset = $positionInGroup * 3;
            $dataCol = chr(66 + $colOffset);

            // Điền dữ liệu vào cột tương ứng (B, E, hoặc H)
            $currentSheet->setCellValue($dataCol . ($groupStartRow + 1), $registration->so_ve_xe ?? '');
            $currentSheet->setCellValue($dataCol . ($groupStartRow + 2), $registration->ho_va_ten ?? '');
            $currentSheet->setCellValue($dataCol . ($groupStartRow + 3), $registration->lop ?? '');
            $currentSheet->setCellValue($dataCol . ($groupStartRow + 4), $registration->bien_so ?? '');
            $currentSheet->setCellValue($dataCol . ($groupStartRow + 5), $registration->loai_xe ?? '');

            $ticketIndex++;
        }

        // Xóa data từ các vé template không dùng trên tất cả sheets
        foreach ($allSheets as $sheetIndex => $currentSheet) {
            // Tìm lại group positions cho sheet hiện tại
            $sheetGroupPositions = [];
            for ($row = 1; $row <= $currentSheet->getHighestRow(); $row++) {
                $value = $currentSheet->getCell('A' . $row)->getValue();
                if ($value === 'VÉ GỬI XE') {
                    $sheetGroupPositions[] = $row;
                }
            }

            $startTicketIndex = $sheetIndex * $maxGroupsPerSheet * 3;
            $endTicketIndex = min($startTicketIndex + ($maxGroupsPerSheet * 3), $totalTickets);

            // Xóa data từ các vé không dùng trong sheet này
            for ($i = $endTicketIndex; $i < $startTicketIndex + (count($sheetGroupPositions) * 3); $i++) {
                $ticketInSheet = $i - $startTicketIndex;
                $groupInSheet = floor($ticketInSheet / 3);
                $positionInGroup = $ticketInSheet % 3;

                if ($groupInSheet >= count($sheetGroupPositions)) {
                    break;
                }

                $groupStartRow = $sheetGroupPositions[$groupInSheet];
                $colOffset = $positionInGroup * 3;
                $dataCol = chr(66 + $colOffset);

                $currentSheet->setCellValue($dataCol . ($groupStartRow + 1), '');
                $currentSheet->setCellValue($dataCol . ($groupStartRow + 2), '');
                $currentSheet->setCellValue($dataCol . ($groupStartRow + 3), '');
                $currentSheet->setCellValue($dataCol . ($groupStartRow + 4), '');
                $currentSheet->setCellValue($dataCol . ($groupStartRow + 5), '');
            }
        }

        // Lưu file
        $fileName = 'Vexe_T' . $this->thang . '_' . $this->nam . '.xlsx';
        $filePath = storage_path('app/exports/' . $fileName);

        // Tạo thư mục exports nếu chưa có
        if (!file_exists(storage_path('app/exports'))) {
            mkdir(storage_path('app/exports'), 0755, true);
        }

        $writer = new Xlsx($spreadsheet);
        $writer->save($filePath);

        // Trả về đường dẫn file để controller tải về
        return $filePath;
    }
}

