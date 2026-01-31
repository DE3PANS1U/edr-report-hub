import os

required_files = [
    'app.py',
    'edr_report_generator_custom.py',
    'requirements.txt',
    'Procfile',
    'templates/index.html',
    'static/style.css',
    'static/script.js',
    'DEPLOYMENT_GUIDE.md'
]

print("Verifying deployment files...")
missing = []
for file in required_files:
    if os.path.exists(file):
        print(f"✅ Found: {file}")
    else:
        print(f"❌ Missing: {file}")
        missing.append(file)

if not missing:
    print("\n🎉 All files ready for deployment!")
else:
    print(f"\n⚠️ Missing {len(missing)} files. Check list above.")
