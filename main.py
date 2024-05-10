"""
Creates a daily excel report from ezee span and demand point layers.
"""
import arcgis
import openpyxl
import pandas
import datetime
from datetime import timezone

def datetime_30_day_ago():
    # Get the current time
    current_time = datetime.datetime.now(timezone.utc)

    # Subtract 15 minutes from the current time
    time_30_days_ago = current_time - datetime.timedelta(days=30)

    return time_30_days_ago

def datetime_7_day_ago():
    # Get the current time
    current_time = datetime.datetime.now(timezone.utc)

    # Subtract 15 minutes from the current time
    time_7_days_ago = current_time - datetime.timedelta(days=7)

    return time_7_days_ago

def datetime_1_day_ago():
    # Get the current time
    current_time = datetime.datetime.now(timezone.utc)

    # Subtract 15 minutes from the current time
    time_1_day_ago = current_time - datetime.timedelta(days=1)

    return time_1_day_ago

def main():
    gis = arcgis.gis.GIS("https://ezeefiber.maps.arcgis.com/home", 'cbrown_lightsett', 'Redcar=1')
    span_id = 'ad7e5f94dd5d46e2a7cab5e5a2c6966a'
    span_item = gis.content.get(span_id)
    span_lyr = span_item.layers[0]
    span_df = span_lyr.query(as_df=True)

    total_length = span_df['CALCULATED_LENGTH'].sum()

    print(total_length)

if __name__ == "__main__":
    main()