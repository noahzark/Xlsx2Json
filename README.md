# Xlsx2Json
A Java parser to convert xlsx sheets to JSON

Supported platforms: Anywhere you can run a Java program

## Quick start
1. Download release or use Gradle to make a build 
2. Run the command line
3. The json file will be generated with the same filename

## Usage

> xlsx2json target_name "sheet_name_1 sheet_name_2 ..." [true|false]

* The first argument is the Excel filename
* The second argument it the sheet you want to export to json
* The third argument indicates whether show the sheet names in generated json or not

(e.g. ture = {"sheet1":{...},"sheet2":{...} | false = [{...},{...}])

## Gradle build command
> $ gradle clean

> $ gradle fatJar

The Jar is created under the ```$project/build/libs/``` folder.

## Example (Excel .xlsx file)
| Integer | String | Basic  | Array\<Double\> | Array\<String\>   | Reference   | Object      |
| ----   | --------| ------ | ---------------- | ---------- | ---------- | ------------ |
| id     | weapon  | flag   | nums  | words  | shiled@shieldStuffs#_id   | objects      |
| 123    | shield  | TRUE   | 1,2   | hello,world   | COPPER_SHIELD | a:123,b:"45",c:false   |
|      | sword   | FALSE  |   | oh god    |   | a:11;b:"22",c:true    |

Result:

```json
{
   "example":[
      {
         "weapon":"shiled",
         "flag":true,
         "objects":{
            "a":123,
            "b":"45",
            "c":false
         },
         "words":[
            "hello",
            "world"
         ],
         "id":123,
         "nums":[
            1,
            2
         ],
         "shiled":{
            "forSale":true,
            "price":3600,
            "durability":20,
            "name":"COPPER SHIELD",
            "priceType":"GOLD",
            "description":"desc",
            "_id":"COPPER_SHIELD",
            "pic":"coppershield.png",
            "life":100
         }
      },
      {
         "weapon":"sword",
         "flag":false,
         "objects":{
            "a":11,
            "b":"22",
            "c":true
         },
         "words":[
            "oh god"
         ],
         "shiled@shieldStuffs#_id":null,
         "id":456,
         "nums":[
            3,
            5,
            8
         ]
      }
   ]
}
```

## Supported types
#### Basic Types
* String
* Integer
* Float
* Double
* Boolean

You can use "Basic" to let the parser automatically detect types

> Especially if all columns are basic types, you can omit the type definition row

#### Array Types
* Array\<String\>
* Array\<Boolean\>
* Array\<Double\>

The values should be divided using commas ","

> You can use the Array<Double> to represent all numeric types like Integer/Float and so on

#### Object type
* Object

Use this one to directly construct a JSON object using basic types, childs should be divided using commas ","

> For more complicated objects, see Reference type

#### Reference type
* Reference

Use this type to insert a JSON object from another sheet, the format should be

``` name_of_this_column@sheet_name#column_name ```

and the value should be the column value of target.

