from optparse import Option

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
        import pdb; pdb.set_trace()

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
