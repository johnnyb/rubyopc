require "open_packaging_conventions/package"
require "open_packaging_conventions/docx_helper"
require "open_packaging_conventions/controller_extensions"

ActionController::Base.send(:include, OpenPackagingConventions::ControllerExtensions)
