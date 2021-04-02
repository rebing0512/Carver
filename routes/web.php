<?php

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

Route::get('/', function () {
    return view('welcome');
});
Route::group([
    'namespace'=>'Api',
],function(\Illuminate\Routing\Router $router){
    # excel导出
    $router->any('phpExcel','ApiController@phpExcel');
    # 初步评审test
    $router->any('test1','ApiController@test1');
    # 唱价记录test
    $router->any('test2','ApiController@test2');
});