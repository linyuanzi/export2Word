<?php

require_once('word/WordAdapter.php');

$host = 'localhost';
$userName = 'root';
$password = 'root';

fwrite(STDOUT, 'Enter the db name:');
$dbName = trim(fgets(STDIN));

fwrite(STDOUT, 'Enter the export document name:');
$name = trim(fgets(STDIN));

try {
	$db = new mysqli($host, $userName, $password, $dbName);
	$db->set_charset('utf8');
	$result = $db->query('SHOW TABLE STATUS');
	$status = $result->fetch_all(MYSQLI_ASSOC);
	foreach ($status as &$v) {
		$table = $v['Name'];
	    $result = $db->query("SHOW FULL FIELDS FROM $table");
	    $v['fields'] = $result->fetch_all(MYSQLI_ASSOC);
	}

	$word = new \WordAdapters();
	$word->exportWord($status, $name);
} catch (Exception $e) {
	die('Error:' . $e->getMessage());
}

return;
