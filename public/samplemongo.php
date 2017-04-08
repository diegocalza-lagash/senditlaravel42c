<?php
/* my_collection */

/* 1 */
{
    "_id" : ObjectId("5707f007639a94cbc600f282"),
    "id" : 1,
    "name" : "Name 1"
}

/* 2 */
{
    "_id" : ObjectId("5707f0a8639a94f4cd2c84b1"),
    "id" : 2,
    "name" : "Name 2"
}


//I'm using the following code:

$filter = ['id' => 2];
$options = [
   'projection' => ['_id' => 0],
];
$query = new MongoDB\Driver\Query($filter, $options);
$rows = $mongo->executeQuery('db_name.my_collection', $query); // $mongo contains the connection object to MongoDB
echo "hola";
foreach($rows as $r){
   print_($r);
}
?>