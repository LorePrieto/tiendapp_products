
lib = File.expand_path("../lib", __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require "tiendapp_products/version"

Gem::Specification.new do |spec|
  spec.name          = "tiendapp_products"
  spec.version       = TiendappProducts::VERSION
  spec.authors       = ["LorePrieto", "FranLucchini"]
  spec.email         = ["mtprieto@uc.cl", "flucchini@uc.cl"]

  spec.summary       = %q{Gem for TiendApp.cl for importing and exporting shop products}
  spec.description   = %q{Gem for TiendApp.cl for importing and exporting shop products}
  spec.homepage      = "https://github.com/LorePrieto/tiendapp_products"

  spec.files         = `git ls-files -z`.split("\x0").reject do |f|
    f.match(%r{^(test|spec|features)/})
  end
  spec.bindir        = "exe"
  spec.executables   = spec.files.grep(%r{^exe/}) { |f| File.basename(f) }
  spec.require_paths = ["lib"]

  spec.add_runtime_dependency 'spree_core', '~> 3.0.4'
  spec.add_runtime_dependency 'axlsx', "~> 1.3.6"
  spec.add_runtime_dependency 'zip-zip'
  spec.add_runtime_dependency "roo", "~> 2.7.1"

  spec.add_development_dependency "bundler", "~> 1.16"
  spec.add_development_dependency "rake", "~> 10.0"
  spec.add_development_dependency 'factory_bot'
  spec.add_development_dependency 'coffee-rails'
  spec.add_development_dependency "rspec"
  spec.add_development_dependency 'rspec-rails'
  spec.add_development_dependency 'database_cleaner'
end
