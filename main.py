# -*- coding: utf-8 -*-
"""Cloud Run entrypoint.

Buildpacks will run: gunicorn -b :$PORT main:app
"""

from app import app
