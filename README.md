<h1 align="center">Mail Security Viewer</h1>

<p align="center">
  <img src="https://github.com/ColdFire1985/EmlAnalyzer/blob/main/main.jpg?raw=true" alt="Main Form Preview" width="50%" />
</p>

<p align="center">
  <strong>A PowerShell-based GUI utility for safe inspection of .eml and .msg files.</strong>
</p>

<hr />

<h2>⚠️ Important Disclaimer</h2>
<p><strong>Use this tool at your own risk.</strong></p>
<ul>
  <li><strong>Registry Modification:</strong> This script modifies the Windows Registry (<code>HKCU</code>) to force the <code>WebBrowser</code> control to use the IE11 rendering engine.</li>
  <li><strong>Data Privacy:</strong> Using the "Upload to VirusTotal" feature sends file data to a third-party service. Do not upload files containing sensitive Personal Identifiable Information (PII).</li>
  <li><strong>No Warranty:</strong> The author is not responsible for any system instability or security flags triggered by the use of this tool.</li>
</ul>

<hr />

<h2>🚀 Features</h2>
<ul>
  <li><strong>Dual Format Support:</strong> Open and parse both Outlook (<code>.msg</code>) and standard (<code>.eml</code>) email files.</li>
  <li><strong>Header Analysis:</strong> Quickly view and copy Message-ID and full transport headers.</li>
  <li><strong>Safe Preview:</strong> Render HTML bodies in a sandboxed environment with a Plain Text fallback.</li>
  <li><strong>Link Extraction:</strong> Automatically identifies and lists all URLs found within the email.</li>
  <li><strong>VirusTotal Integration:</strong> Extract attachments or upload them directly for reputation analysis.</li>
</ul>

<hr />

<h2>🛠️ Requirements</h2>
<ul>
  <li><strong>OS:</strong> Windows 10/11</li>
  <li><strong>PowerShell:</strong> 5.1 or Desktop Core</li>
  <li><strong>Dependencies:</strong> <code>System.Windows.Forms</code> & <code>System.Drawing</code> (Standard Windows libraries).</li>
</ul>

<hr />

<h2>🔧 Installation & Setup</h2>

<h3>1. Clone the repository</h3>
<pre><code>git clone https://github.com/ColdFire1985/EmlAnalyzer.git</code></pre>

<h3>2. Configure API Key</h3>
<p>Open the script and locate the <code>$VT_API_KEY</code> variable. It is recommended to use an environment variable:</p>
<pre><code>$VT_API_KEY = $env:VT_API_KEY</code></pre>

<h3>3. Run the script</h3>
<p>Right-click <code>Viewer.ps1</code> and select <strong>Run with PowerShell</strong>.</p>

<hr />

<h2>📄 License</h2>
<p>This project is licensed under the MIT License.</p>
