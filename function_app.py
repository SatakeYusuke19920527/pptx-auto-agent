import azure.functions as func  # type: ignore
from blueprints.create_pptx import create_pptx_bp  # type: ignore

app = func.FunctionApp()

# Blueprintsの登録
app.register_blueprint(create_pptx_bp)
