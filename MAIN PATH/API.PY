import Monitoring
import json
from flask import Flask
from SaraCalendar import main


mn = main()
datas = Monitoring.get_data()
app = Flask(__name__)
@app.route('/')
def index():
    return json.dumps({"datas": datas, "ESCALAS ÁTRIO MUSIC": mn})
app.run()