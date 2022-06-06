# Dj easy xls

Helpful class and methods to import and export xls with django framework.

## Usage

### Import from xls

We call method get_sheet_rows which converts the table into dict with column as keys.

``` py
excel = OpenpyxlImport(file)
rows = excel.get_sheet_rows()
if excel.tally_header(rows[0], self.fields):
    for row in rows[1:]:
        params = excel.row_to_dict(row)
        print(params)
```

### Export from xls

Simple example to export django model queryset into csv file.

```py
file_name = self.file_name + ' ' + datetime.datetime.today().strftime('%Y-%m-%d') or 'Untitled'
export = OpenpyxlExport(file_name)
export.generate(self.fields, True)
for object in self.queryset:
    values = [change_format(object, val) for val in self.fields]
    export.generate(values)
export.set_width() # sets proper width of each columns
```

#### Return xlsx file as a response

Once we generate the xls in an export instance we can return response as

```py
return export.response()
```

#### Return xlsx file as a response

Saving xlsx in a directory path.

```py
return export.wb.save("<path>/test.xlsx")
```

### Saving response from django as a file with axios
Django http response can be saved as a file from an axios request
```es6
const url = '/download'
const config = {
  baseURL: process.env.BaseURL,
  responseType: "blob", // Very important!
};
try {
  const response = await axios.get(url, config);
  const url = window.URL.createObjectURL(new Blob([response.data]));
  const link = document.createElement("a");
  link.href = url;
  link.setAttribute("download", "file.xlsx"); //or any other extension
  document.body.appendChild(link);
  link.click();
} catch (error) {
  console.log(error);
}
```
