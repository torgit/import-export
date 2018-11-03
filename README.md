## Installation

```bash
$ npm install
```

## Running the app

```bash
# development
$ npm run start
```

## Example export xlsx

```bash
curl -X POST \
  http://localhost:3000/export/excel \
  -H 'Content-Type: application/json' \
  -H 'cache-control: no-cache' \
  -d '{
	"data": [
		{
			"name": "dog1",
			"breed": "dog",
			"age": 2,
			"origin": {
				"country": "TH",
				"city": "BKK"
			}
		},
		{
			"name": "bird1",
			"breed": "bird",
			"age": 1,
			"origin": {"city": "BKK"}
		},
		{
			"name": "cat1",
			"breed": "cat",
			"age": 7,
			"origin": {
				"country": "TH",
				"city": "BKK"
			}
		},
		{
			"name": "bird1",
			"breed": "bird",
			"age": 1,
			"origin": {
				"country": "TH"
			}
		}
	]
}'
```

## Example export json

```bash
curl -X POST \
  http://localhost:3000/export/json \
  -H 'Content-Type: application/x-www-form-urlencoded' \
  -H 'cache-control: no-cache' \
  -H 'content-type: multipart/form-data; boundary=----WebKitFormBoundary7MA4YWxkTrZu0gW' \
  -F file=@/home/import-export/result.xlsx
```
