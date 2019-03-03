from nrcdataproxy.storage import IncidentStore
import pymongo

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
