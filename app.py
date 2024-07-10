from flask import Flask, render_template,request,jsonify
# import pandas as pd
# import matplotlib.pyplot as plt
# import numpy as np
# from datetime import datetime
# from matplotlib import patches 
# from matplotlib.patches import Wedge
# from io import BytesIO
# import base64
# import matplotlib.pyplot as plt
# from matplotlib.gridspec import GridSpec
# import plotly.graph_objs as go
# import plotly.io as pio
# from flask_cors import CORS

#file_path = 'https://bpxai-my.sharepoint.com/personal/manas_shalgar_bpx_ai/_layouts/15/download.aspx?share=EW7IZR3VH_ZDp_rhgbv9GlQBowPIosXVpfUCXsAUTYNJ8Q'


app = Flask(__name__, static_folder='static')

@app.route('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':

    app.run(debug=True)