from flask import Flask
from flask_cors import CORS
from flask_restx import Api, Resource, fields
import pandas as pd

app = Flask(__name__)
api = Api(app, version='1.0', title='Excel API',
          description='A simple API for manipulating Excel files')
CORS(app, supports_credentials=True, origins=['http://localhost:3000'])

EXCEL_FILE = 'data.xlsx'  # Define the Excel file path


ns = api.namespace('rows', description='Operations related to Excel rows')

row_model = api.model('Row', {
    'column1': fields.String(required=True, description='Column 1 value'),
    'column2': fields.String(required=True, description='Column 2 value'),
    # Define other columns as needed
})

@ns.route('/')
class RowsList(Resource):
    @ns.doc('list_rows')
    def get(self):
        """List all rows"""
        df = pd.read_excel(EXCEL_FILE, engine=None)
        return df.to_dict()

    @ns.doc('create_row')
    @ns.expect(row_model)
    def post(self):
        """Add a new row"""
        df = pd.read_excel(EXCEL_FILE, engine=None)
        new_row = pd.DataFrame(api.payload, index=[0])
        df = pd.concat([df, new_row], ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        return {'message': 'Row added.'}, 201

@ns.route('/<int:index>')
@ns.param('index', 'The row index')
@ns.response(404, 'Row not found')
class Row(Resource):
    @ns.doc('get_row')
    def get(self, index):
        """Fetch a single row"""
        df = pd.read_excel(EXCEL_FILE)
        if index >= len(df):
            ns.abort(404)
        return df.iloc[index].to_dict()

    @ns.doc('delete_row')
    def delete(self, index):
        """Delete a row"""
        df = pd.read_excel(EXCEL_FILE, engine=None)
        if index >= len(df):
            ns.abort(404)
        df = df.drop(index).reset_index(drop=True)
        df.to_excel(EXCEL_FILE, index=False)
        return {'message': 'Row deleted.'}, 204

    @ns.doc('update_row')
    @ns.expect(row_model)
    def put(self, index):
        """Update a row"""
        df = pd.read_excel(EXCEL_FILE, engine=None)
        if index >= len(df):
            ns.abort(404)
        for key, value in api.payload.items():
            df.at[index, key] = value
        df.to_excel(EXCEL_FILE, index=False)
        return {'message': 'Row updated.'}, 200

if __name__ == '__main__':
    app.run(debug=True)
