<?php

use App\Http\Controllers\VehicleRegistrationController;
use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
*/

// Vehicle Registration Routes
Route::get('/', [VehicleRegistrationController::class, 'index'])->name('vehicle-registrations.index');
Route::post('/vehicle-registrations/import', [VehicleRegistrationController::class, 'import'])->name('vehicle-registrations.import');
Route::get('/vehicle-registrations/available-months', [VehicleRegistrationController::class, 'getAvailableMonths'])->name('vehicle-registrations.available-months');
Route::post('/vehicle-registrations/export', [VehicleRegistrationController::class, 'export'])->name('vehicle-registrations.export');
Route::delete('/vehicle-registrations/delete-all', [VehicleRegistrationController::class, 'deleteAll'])->name('vehicle-registrations.delete-all');
