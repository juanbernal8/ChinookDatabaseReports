import pandas as pd
import sqlite3
import numpy as np

##### INPUT #####

## Conecting to the database
db_connection = sqlite3.connect(
    'C:/Files/chinook.db'
    )

#### Querying the database Using Pandas read_sql function to get dataframes

## Count of customers per country
df_customers_per_coutry = pd.read_sql(
    "SELECT customers.Country, count(customers.CustomerId) as CustomersCount\
    FROM customers \
    GROUP BY customers.Country \
    ORDER BY 2 DESC, 1",
    db_connection
    )

## Top 100 songs by sales
df_100_songs_by_sales = pd.read_sql(
    "SELECT tracks.Name as 'Song', artists.Name as 'Artist', albums.Title as 'Album', sum(invoices.total) as TotalSales \
    FROM invoices \
    INNER JOIN invoice_items \
        ON invoices.InvoiceId = invoice_items.InvoiceId \
    INNER JOIN tracks \
        ON invoice_items.TrackId = tracks.TrackId \
    INNER JOIN albums  \
        ON tracks.AlbumId = albums.AlbumId \
    INNER JOIN artists \
        ON albums.ArtistId = artists.ArtistId \
    GROUP BY invoice_items.TrackId \
    ORDER BY TotalSales DESC \
    LIMIT 100",
    db_connection
    )
    
## Price of the entire collection of songs for each Artist
## Only for those artists that have more than 5 songs
df_entire_collection_artist_price = pd.read_sql(
    "SELECT artists.Name as 'ArtistName', count(tracks.TrackId) as 'TotalSongs' ,sum(tracks.UnitPrice) as 'TotalPrice' \
    FROM tracks \
    INNER JOIN albums \
        ON tracks.AlbumId = albums.AlbumId \
    INNER JOIN artists \
        ON albums.ArtistId = artists.ArtistId \
    GROUP BY artists.name \
    HAVING TotalSongs > 5 \
    ORDER BY TotalPrice DESC",
    db_connection
    )
    
## Songs by genre
## Importing and adding an aditional column for future plotting
df_songs_by_genre = pd.read_sql(
    "SELECT genres.Name, count(tracks.TrackId) as 'Songs' \
    FROM tracks \
    INNER JOIN genres \
        ON tracks.GenreId = genres.GenreId \
    GROUP BY genres.Name \
    ORDER BY Songs DESC",
    db_connection
    )
df_songs_by_genre['Running Total %'] = \
    (df_songs_by_genre['Songs'] / sum(df_songs_by_genre['Songs'])).cumsum()


##### OUTPUT #####

##Creating writer variable
writer = pd.ExcelWriter('C:/Files/chinook_reports.xlsx')

## Writing dataframes to excel sheets
df_customers_per_coutry.set_index(
    np.arange(1,len(df_customers_per_coutry['Country'])+1)).to_excel(
        writer,
        sheet_name='CustomersPerCountry',
        engine='xlsxwritter'
        )

df_100_songs_by_sales.set_index(
    np.arange(1,len(df_100_songs_by_sales['Song'])+1)).to_excel(
        writer,
        sheet_name='100SongsBySales',
        engine='xlsxwritter'
        )

df_entire_collection_artist_price.to_excel(
    writer,
    sheet_name='EntireCollArtistPrice',
    index=False,
    engine='xlsxwritter'
    )

df_songs_by_genre.set_index(
    np.arange(1,(len(df_songs_by_genre['Songs'])+1))).to_excel(
        writer,
        sheet_name='SongsByGenre',
        engine='xlsxwritter'
        )


## Creating the workbook and worksheets for each report
workbook = writer.book
worksheet_customers_per_coutry = writer.sheets['CustomersPerCountry']
worksheet_100_songs_by_sales = writer.sheets['100SongsBySales']
worksheet_entire_collection_artist_price = writer.sheets['EntireCollArtistPrice']
worksheet_songs_by_genre = writer.sheets['SongsByGenre']


## Creating center format for cells
center = workbook.add_format({
    'align':'center'
    })

## Creating currency format for specific cells
currency = workbook.add_format({
    'align': 'center',
    'num_format': '$#,##0.00'
    })

## Creating percentage format for specific cells
percentage = workbook.add_format({
    'align': 'center',
    'num_format': '0.00%'
    })
    

## Setting column style for each sheet
worksheet_customers_per_coutry.set_column(
    'B:C',
    18,
    center
    )

worksheet_100_songs_by_sales.set_column(
    'B:D',
    26,
    center
    )

worksheet_100_songs_by_sales.set_column(
    'E:E',
    None,
    currency,
    )

worksheet_entire_collection_artist_price.set_column(
    'A:A',
    35,
    center
    )

worksheet_entire_collection_artist_price.set_column(
    'B:B',
    None,
    center
    )

worksheet_entire_collection_artist_price.set_column(
    'C:C',
    None,
    currency
    )

worksheet_entire_collection_artist_price.set_column(
    'A:A',
    26,
    center
    )

worksheet_songs_by_genre.set_column(
    'B:C',
    17,
    center
    )

worksheet_songs_by_genre.set_column(
    'D:D',
    14,
    percentage
    )


#### Adding a pareto chart
## Creating the primary chart
column_chart = workbook.add_chart({
    'type': 'column'
    })

column_chart.set_size({
     'width': 885,
     'height': 500
     })

## Configuring the data series for the primary chart.
column_chart.add_series({
    'name':       'Songs Count',
    'categories': '=SongsByGenre!B2:B25',
    'values':     '=SongsByGenre!C2:C25'
    })

## Creating the secondary chart
line_chart = workbook.add_chart({
    'type': 'line'
    })

line_chart.set_size({
     'width': 885,
     'height': 500
     })

## Configuring the data series for the secondary chart.
line_chart.add_series({
    'name':       'Running Total %',
    'categories': '=SongsByGenre!B2:B25',
    'values':     '=SongsByGenre!D2:D25',
    'y2_axis': True
    })

## Combining the charts
column_chart.combine(line_chart)

## Adding Title and labels to all the axis
column_chart.set_title({  'name': 'Songs By Genre - Pareto Chart'})
column_chart.set_x_axis({ 'name': 'Genre'})
column_chart.set_y_axis({ 'name': 'Songs Count'})
## Note: The y2 properties are on the secondary chart.
line_chart.set_y2_axis({'name': 'Running Percentage'})

## Inserting the chart into the worksheet
worksheet_songs_by_genre.insert_chart(
    'E2',
    column_chart
    )

## Saving the file
writer.save()

