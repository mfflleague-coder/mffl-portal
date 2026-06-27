<?php

$UPLOAD_KEY = "MFFL2026_UPLOAD_9X7K2P";

if ($_SERVER["REQUEST_METHOD"] !== "POST") {
    http_response_code(405);
    echo "Method not allowed";
    exit;
}

$key = isset($_POST["key"]) ? $_POST["key"] : "";
$playerId = isset($_POST["playerId"]) ? $_POST["playerId"] : "";

if ($key !== $UPLOAD_KEY) {
    http_response_code(403);
    echo "Forbidden";
    exit;
}

if ($playerId === "" || !isset($_FILES["photo"])) {
    http_response_code(400);
    echo "Missing playerId or photo";
    exit;
}

$playerId = preg_replace("/[^A-Za-z0-9_-]/", "", $playerId);

$targetDir = __DIR__ . "/assets/players/";

if (!is_dir($targetDir)) {
    mkdir($targetDir, 0755, true);
}

$targetFile = $targetDir . $playerId . ".jpg";

$tmpFile = $_FILES["photo"]["tmp_name"];

$imageInfo = getimagesize($tmpFile);

if ($imageInfo === false) {
    http_response_code(400);
    echo "Invalid image";
    exit;
}

if (!move_uploaded_file($tmpFile, $targetFile)) {
    http_response_code(500);
    echo "Upload failed";
    exit;
}

header("Content-Type: application/json");

echo json_encode(array(
    "success" => true,
    "url" => "https://portal.mffl.co.uk/assets/players/" . $playerId . ".jpg"
));

?>