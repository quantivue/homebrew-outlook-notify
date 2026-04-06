class OutlookNotify < Formula
  include Language::Python::Virtualenv

  desc "Menu bar notifier for Outlook subfolder emails on macOS"
  homepage "https://github.com/quantivue/outlook-notify"
  # UPDATE sha256 after publishing the first GitHub release:
  #   shasum -a 256 /path/to/downloaded.tar.gz
  url "https://github.com/quantivue/outlook-notify/archive/refs/tags/v1.0.0.tar.gz"
  sha256 "PLACEHOLDER_UPDATE_AFTER_FIRST_RELEASE"
  license "MIT"

  depends_on "python@3.13"
  depends_on :macos

  # pip install rumps pyobjc-core pyobjc-framework-Cocoa
  # shasum -a 256 on the .tar.gz from PyPI

  resource "pyobjc-core" do
    url "https://files.pythonhosted.org/packages/source/p/pyobjc-core/pyobjc_core-12.1.tar.gz"
    sha256 "2bb3903f5387f72422145e1466b3ac3f7f0ef2e9960afa9bcd8961c5cbf8bd21"
  end

  resource "pyobjc-framework-Cocoa" do
    url "https://files.pythonhosted.org/packages/source/p/pyobjc-framework-Cocoa/pyobjc_framework_cocoa-12.1.tar.gz"
    sha256 "5556c87db95711b985d5efdaaf01c917ddd41d148b1e52a0c66b1a2e2c5c1640"
  end

  resource "rumps" do
    url "https://files.pythonhosted.org/packages/source/r/rumps/rumps-0.4.0.tar.gz"
    sha256 "17fb33c21b54b1e25db0d71d1d793dc19dc3c0b7d8c79dc6d833d0cffc8b1596"
  end

  def install
    virtualenv_install_with_resources
    bin.install "outlook-notify.py" => "outlook-notify"
  end

  service do
    run [opt_bin/"outlook-notify"]
    keep_alive true
    log_path var/"log/outlook-notify.log"
    error_log_path var/"log/outlook-notify.err"
  end

  def caveats
    <<~EOS
      After installing, select folders to watch from the 📬 menu bar icon.

      Outlook must be running for notifications to work (it handles Exchange sync).

      Manage the service:
        brew services start outlook-notify
        brew services stop  outlook-notify
    EOS
  end

  test do
    assert_predicate bin/"outlook-notify", :exist?
    system bin/"outlook-notify", "--help" rescue nil
  end
end
