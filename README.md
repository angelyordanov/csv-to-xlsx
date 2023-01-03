#### csv-to-xlsx

```sh
export CSV=/Users/angel/Downloads/2022-07.csv

# replace &lt;br\/&gt;
sed -i '' 's/&lt;br\/&gt;/ /g' $CSV

# reverse the lines
sed -i '' '1!G;h;$!d' $CSV

# move the last line to the top
sed -i '' '1h;1d;$!H;$!d;G' $CSV

# convert to xlsx
dotnet run -d I -d J -ds "," -dt A -dtf "dd.MM.yyyy" $CSV
```
