class OutlookNotify < Formula
  include Language::Python::Virtualenv

  desc "Menu bar notifier for Outlook subfolder emails on macOS"
  homepage "https://github.com/quantivue/homebrew-outlook-notify"
  url "https://github.com/quantivue/homebrew-outlook-notify/archive/refs/tags/v1.0.0.tar.gz"
  sha256 "453d059b931ceec86e6ad00c64ab4989668ff354a7e6ea93d55336b70aadabb9"
  license "MIT"

  depends_on "python@3.13"
  depends_on :macos

  # Pre-compiled universal2 wheels — no Xcode required
  resource "pyobjc-core" do
    url "https://files.pythonhosted.org/packages/cp313/p/pyobjc_core/pyobjc_core-12.1-cp313-cp313-macosx_10_13_universal2.whl"
    sha256 "01c0cf500596f03e21c23aef9b5f326b9fb1f8f118cf0d8b66749b6cf4cbb37a"
  end

  resource "pyobjc-framework-Cocoa" do
    url "https://files.pythonhosted.org/packages/cp313/p/pyobjc_framework_cocoa/pyobjc_framework_cocoa-12.1-cp313-cp313-macosx_10_13_universal2.whl"
    sha256 "5a3dcd491cacc2f5a197142b3c556d8aafa3963011110102a093349017705118"
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
