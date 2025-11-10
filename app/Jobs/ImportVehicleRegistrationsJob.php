<?php

namespace App\Jobs;

use App\Imports\VehicleRegistrationsImport;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use Maatwebsite\Excel\Facades\Excel;

class ImportVehicleRegistrationsJob implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    protected $filePath;

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
    public function __construct($filePath)
    {
        $this->filePath = $filePath;
    }

    /**
     * Execute the job.
     */
    public function handle(): void
    {
        try {
            // Import file Excel với chunk reading
            Excel::import(new VehicleRegistrationsImport, $this->filePath);

            // Xóa file sau khi import thành công (tùy chọn)
            // if (file_exists($this->filePath)) {
            //     unlink($this->filePath);
            // }

        } catch (\Exception $e) {
            // Log lỗi
            \Log::error('Import failed: ' . $e->getMessage());

            // Throw lại exception để job có thể retry
            throw $e;
        }
    }

    /**
     * Xử lý khi job thất bại
     */
    public function failed(\Throwable $exception): void
    {
        // Log lỗi khi job thất bại hoàn toàn
        \Log::error('Import job failed completely: ' . $exception->getMessage());
    }
}
