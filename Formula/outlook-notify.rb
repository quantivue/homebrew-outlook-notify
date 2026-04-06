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
    url "https://github.com/quantivue/homebrew-outlook-notify/releases/download/v1.0.0/rumps-0.4.0-py3-none-any.whl"
    sha256 "4da62e8598d99f84facf4d0a509dbda58b3484fcda2a88149c61ea90850c2d90"
  end

  def install
    venv = virtualenv_create(libexec, "python@3.13")

    # Install each resource — handles both .whl files and source packages
    resources.each do |r|
      r.stage do
        whl = Dir["*.whl"].first
        if whl
          system libexec/"bin/pip", "install", "--no-deps", whl
        else
          system libexec/"bin/pip", "install", "--no-deps", "."
        end
      end
    end

    # Install the script and create a wrapper that uses the virtualenv Python
    pkgshare.install "outlook-notify.py"
    (bin/"outlook-notify").write <<~SH
      #!/bin/bash
      exec "#{libexec}/bin/python3" "#{pkgshare}/outlook-notify.py" "$@"
    SH
    chmod 0755, bin/"outlook-notify"
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
