from optparse import Option, OptionParser
from flask import Flask, request
import flask
import json

webapp = Flask(__name__)

@webapp.route('/incidents/<seqnos>', methods=['PUT'])
def save(seqnos):
    print "saving %d" % seqnos

def serve():

    parser = OptionParser(usage="usage: %%prog [options]\n%s" % serve.__doc__)
    parser.add_option(Option("--host", dest="host", default="0.0.0.0"))
    parser.add_option(Option("--port", dest="port", default=5001, type="int"))
    parser.add_option(Option("--debug", dest="debug", default=False, action="store_true"))

    from nrcdataproxy.storage.mongo import MongoIncidentStore
    for option in MongoIncidentStore.commandline_options():
        parser.add_option(option)

    (options, args) = parser.parse_args()

    store = MongoIncidentStore.configure_from_commandline(options)

    webapp.config.store = store
    webapp.run(host=options.host,
               port=options.port,
               debug=options.debug)
