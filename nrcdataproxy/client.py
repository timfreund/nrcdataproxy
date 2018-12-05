import json
import requests

class NRCDataClient:
    def __init__(self, url):
        self.url = url
        self.default_headers = {'Content-type': 'application/json'}

    def save(self, incident):
        r = requests.put("%(url)s/incidents/%(incident_id)d" % {'incident_id':incident['seqnos'],
                                                                'url': self.url},
                         data=json.dumps(incident),
                         headers = self.default_headers)
        print("%d: %s" % (incident['seqnos'], r.status_code))

