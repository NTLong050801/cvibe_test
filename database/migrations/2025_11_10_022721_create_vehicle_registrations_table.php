<?php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
        Schema::create('vehicle_registrations', function (Blueprint $table) {
            $table->id();
            $table->integer('stt')->nullable()->comment('Số thứ tự');
            $table->string('ho_va_ten')->comment('Họ và tên');
            $table->string('lop')->comment('Lớp');
            $table->string('loai_xe')->comment('Loại xe (Xe điện, Xe máy, v.v.)');
            $table->string('bien_so')->nullable()->comment('Biển số xe');
            $table->string('so_ve_xe')->unique()->comment('Số vé xe (TC.001, TC.002, v.v.)');
            $table->date('ngay_dang_ky')->nullable()->comment('Ngày đăng ký');
            $table->integer('thang')->comment('Tháng');
            $table->integer('nam')->comment('Năm');
            $table->timestamps();

            // Index để query nhanh theo tháng năm
            $table->index(['thang', 'nam']);
        });
    }

    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('vehicle_registrations');
    }
};
