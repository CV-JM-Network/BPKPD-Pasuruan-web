<?php
function getMutasiPBB($idmutasi_pbb) {
    $url = "http://api.pendapatan.pasuruankab.go.id/api/get/mutasi/pbb?idmutasi_pbb=" . urlencode($idmutasi_pbb);
    $token = "8853147364"; // Bearer Token

    // Inisialisasi cURL
    $ch = curl_init();

    // Konfigurasi cURL
    curl_setopt($ch, CURLOPT_URL, $url);
    curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
    curl_setopt($ch, CURLOPT_HTTPHEADER, [
        "Authorization: Bearer $token",
        "Accept: application/json"
    ]);

    // Eksekusi request
    $response = curl_exec($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);

    // Cek jika terjadi error
    if (curl_errno($ch)) {
        return json_encode(["status" => "error", "message" => curl_error($ch)]);
    }

    // Tutup koneksi cURL
    curl_close($ch);

    // Return hasil dalam bentuk JSON
    return json_encode([
        "status" => $httpCode == 200 ? "success" : "error",
        "http_code" => $httpCode,
        "data" => json_decode($response, true)
    ], JSON_PRETTY_PRINT);
}

// Contoh penggunaan
$idmutasi_pbb = 4; // Bisa diganti dengan ID lain
header("Content-Type: application/json");
echo getMutasiPBB($idmutasi_pbb);
