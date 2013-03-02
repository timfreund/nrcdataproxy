from datetime import datetime, timedelta
from optparse import Option
import pymongo

class IncidentStore():
    def __init__(self, **kw):
        pass

    @classmethod
    def configuration_options(klass):
        return {}

    @classmethod
    def commandline_options(klass):
        options = []
        for k, v in klass.configuration_options().items():
            options.append(Option(None, "--%s" % k, default=v, dest=k))
        return options

    @classmethod
    def configure_from_commandline(klass, options):
        klass_options = klass.configuration_options().keys()
        kwargs = {}
        for option in dir(options):
            if option in klass_options:
                kwargs[option] = getattr(options, option)

        return klass(**kwargs)

    def save(self, incident):
        pass

    def get(self, incident_id):
        pass

    def find(self, **criteria):
        pass

class MuxingStore(IncidentStore):
    def __init__(self, stores, **kw):
        self.stores = stores

    def save(self, incident):
        for store in self.stores:
            store.save(incident)

class DebugStore(IncidentStore):
    def __init__(self, log=None, **kw):
        if log is None:
            import sys
            log = sys.stdout
    
    def save(self, incident):
        log.write("Pretending to save %d\n" % incident['SEQNOS'])

class MongoIncidentStore(IncidentStore):
    def __init__(self, mongohost='localhost', mongoport=27017, database='nrc', **kw):
        self.host = mongohost
        self.port = mongoport
        self.database_name = database
        self.mongo_server = pymongo.Connection(self.host, self.port)
        self.mongodb = self.mongo_server[self.database_name]
        self.collection = self.mongodb.incidents

    @classmethod
    def configuration_options(klass):
        return {'mongohost': 'localhost',
                'mongoport': 27017,
                'database': 'nrc'}

    def save(self, incident):
        r = {}
        for k, v in incident.items():
            value = v
            if k.count('date_') > 0:
                if v != '' and v is not None and not isinstance(v, datetime):
                    value = datetime.strptime(v, '%Y-%m-%dT%H:%M:%S')
            elif v == u'':
                value = None
            elif type(v) == dict:
                value = {}
                for nested_k, nested_v in v.items():
                    value[nested_k.lower()] = nested_v 

            r[k.lower()] = value
                
        r['_id'] = int(incident['seqnos'])
        self.collection.save(r)

    def find(self, **kw):
        self.collection.find(**kw)
