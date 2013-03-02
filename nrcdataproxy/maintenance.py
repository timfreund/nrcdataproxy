from optparse import Option, OptionParser


# Geocoder
# 1a. Find all objects where geoloc && geoloc_error == None and have a lat_deg value
# 2a. Convert all objects' component values into a geoloc value

# 1b. Find all objects where geoloc && geoloc_error == None and have no lat_deg value
# 2b. look for location_[address, nearest_city, state, zip].  Missing components?  record the eror in geoloc_error
# 3b. Geocode the address.  Record errors in geoloc_error, save the values to geoloc if successful
# ... will need to define geocoder implemenations and geocoding limits per day per IP/authtoken

def geo_convert_dms_to_decimal(degrees, quadrant, minutes, seconds):
    if seconds is None:
        seconds = 0

    hemisphere = 1
    if quadrant.upper() in ['S', 'W']:
        hemisphere = -1

    return (degrees + (minutes / 60.0) + (seconds / 3600.0)) * hemisphere

# 963678: 29N 16' 31" 88W 44' 31"
# ... in google:  29 16' 31", -88 44' 31"
# Thanks to http://dtbaker.com.au/random-bits/google-maps-longitude---latitude---degrees---minutes---seconds-search.html
# and http://www.mikestechblog.com/joomla/misc/gps-category-blog/93-how-to-enter-latitude-longitude-gps-coordinates-google-maps-article.html


def geocode_command():
    parser = OptionParser(usage="usage: %%prog [options]\n%s" % geocode_command.__doc__)
    from nrcdataproxy.storage.mongo import MongoIncidentStore
    for option in MongoIncidentStore.commandline_options():
        parser.add_option(option)

    (options, args) = parser.parse_args()

    store = MongoIncidentStore.configure_from_commandline(options)
    cursor = store.collection.find({'geoloc': None, 'geoloc_error': None, 'lat_deg': {'$ne': None}})
    # ... need to redo things?  Try this one...
    # cursor = store.collection.find({'geoloc': {'$ne': None}, 'lat_deg': {'$ne': None}})

    for doc in cursor:
        try:
            latitude = geo_convert_dms_to_decimal(int(doc['lat_deg']),
                                                  doc['lat_quad'],
                                                  int(doc['lat_min']),
                                                  int(doc['lat_sec']))
            longitude = geo_convert_dms_to_decimal(int(doc['long_deg']),
                                                   doc['long_quad'],
                                                   int(doc['long_min']),
                                                   int(doc['long_sec']))
            doc['geoloc'] = [latitude, longitude]
            print "%d: %f, %f" % (doc['seqnos'], latitude, longitude)

        except Exception, e:
            print "%d: %s" % (doc['seqnos'], e.message)
            doc['geoloc_error'] = e.message

        store.save(doc)




    

