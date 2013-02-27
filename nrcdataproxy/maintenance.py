

# Geocoder
# 1a. Find all objects where geoloc && geoloc_error == None and have a lat_deg value
# 2a. Convert all objects' component values into a geoloc value

# 1b. Find all objects where geoloc && geoloc_error == None and have no lat_deg value
# 2b. look for location_[address, nearest_city, state, zip].  Missing components?  record the eror in geoloc_error
# 3b. Geocode the address.  Record errors in geoloc_error, save the values to geoloc if successful
# ... will need to define geocoder implemenations and geocoding limits per day per IP/authtoken

def geo_convert_dms_to_decimal(degrees, minutes, seconds):
    if isinstance(minutes, basestring):
        minutes = int(minutes)

    if isinstance(seconds, basestring):
        seconds = int(seconds)

    hemisphere = 1
    if isinstance(degrees, basestring):
        degrees = degrees.upper()
        for direction in ['S', 'W']:
            if degrees.count(direction):
                hemisphere = -1
                degrees = degrees.replace(direction, '')
        for direction in ['N', 'E']:
            degrees = degrees.replace(direction, '')
        degrees = int(degrees)

    return (degrees + (minutes / 60.0) + (seconds / 3600.0)) * hemisphere
    
            

        

# 963678: 29N 16' 31" 88W 44' 31"
# ... in google:  29 16' 31", -88 44' 31"
# Thanks to http://dtbaker.com.au/random-bits/google-maps-longitude---latitude---degrees---minutes---seconds-search.html
# and http://www.mikestechblog.com/joomla/misc/gps-category-blog/93-how-to-enter-latitude-longitude-gps-coordinates-google-maps-article.html
