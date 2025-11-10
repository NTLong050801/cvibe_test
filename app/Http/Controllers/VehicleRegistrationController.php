<?php

namespace App\Http\Controllers;

use App\Models\VehicleRegistration;
use Illuminate\Http\Request;

class VehicleRegistrationController extends Controller
{
    public function index(Request $request)
    {
        $query = VehicleRegistration::query();

        // Lọc theo tháng năm nếu có
        if ($request->has('thang') && $request->thang) {
            $query->where('thang', $request->thang);
        }
        if ($request->has('nam') && $request->nam) {
            $query->where('nam', $request->nam);
        }

        // Phân trang
        $registrations = $query->orderBy('created_at', 'desc')->paginate(15);

        return view('vehicle-registrations.index', compact('registrations'));
    }

    public function import(Request $request)
    {
        $request->validate([
            'file' => 'required|mimes:xlsx,xls|max:10240', // Max 10MB
        ]);

        try {
            $file = $request->file('file');
            $fileName = time() . '_' . $file->getClientOriginalName();
            $filePath = $file->storeAs('imports', $fileName);

            // Dispatch job để xử lý import
            \App\Jobs\ImportVehicleRegistrationsJob::dispatch(storage_path('app/' . $filePath));

            return back()->with('success', 'File đã được upload! Đang xử lý import trong background...');
        } catch (\Exception $e) {
            return back()->with('error', 'Có lỗi xảy ra: ' . $e->getMessage());
        }
    }

    public function getAvailableMonths()
    {
        $data = VehicleRegistration::selectRaw('DISTINCT nam, thang')
            ->orderBy('nam', 'desc')
            ->orderBy('thang', 'desc')
            ->get()
            ->groupBy('nam')
            ->map(function ($items) {
                return $items->pluck('thang')->toArray();
            });

        return response()->json($data);
    }

    public function export(Request $request)
    {
        $request->validate([
            'thang' => 'required|integer|min:1|max:12',
            'nam' => 'required|integer|min:2020|max:2030',
        ]);

        $thang = $request->thang;
        $nam = $request->nam;

        // Kiểm tra xem có dữ liệu không
        $count = VehicleRegistration::where('thang', $thang)
            ->where('nam', $nam)
            ->count();

        if ($count === 0) {
            return back()->with('error', "Không có dữ liệu cho tháng {$thang}/{$nam}!");
        }

        try {
            // Gọi Job để xử lý export
            $job = new \App\Jobs\ExportVehicleRegistrationsJob($thang, $nam);
            $filePath = $job->handle();

            // Kiểm tra file có tồn tại không
            if (!file_exists($filePath)) {
                return back()->with('error', 'Có lỗi xảy ra khi tạo file export!');
            }

            // Tải file về
            $fileName = basename($filePath);
            return response()->download($filePath, $fileName, [
                'Content-Type' => 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            ])->deleteFileAfterSend(true);

        } catch (\Exception $e) {
            return back()->with('error', 'Có lỗi xảy ra khi export: ' . $e->getMessage());
        }
    }

    public function deleteAll()
    {
        try {
            $count = VehicleRegistration::count();
            VehicleRegistration::truncate();

            return redirect()->route('vehicle-registrations.index')
                ->with('success', "Đã xóa thành công {$count} bản ghi!");
        } catch (\Exception $e) {
            return back()->with('error', 'Có lỗi xảy ra khi xóa dữ liệu: ' . $e->getMessage());
        }
    }
}
