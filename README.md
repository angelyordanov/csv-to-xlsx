#### csv-to-xlsx

1. replace `&lt;br\/&gt;` with ` `

```sh
sed -i '' 's/&lt;br\/&gt;/ /g' /Users/angel/Downloads/2022-07.csv
```

2. convert to xlsx

```sh
dotnet run -d I -d J -ds "," -dt A -dtf "dd.MM.yyyy" "/Users/angel/Downloads/2022-07.csv"
```
