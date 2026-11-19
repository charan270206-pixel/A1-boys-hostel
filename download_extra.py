import urllib.request
import os

extra_images = {
    'slide4_phishing.jpg': 'https://images.unsplash.com/photo-1563089145-599997674d42?w=600&auto=format&fit=crop&q=80',
    'slide4_ransomware.jpg': 'https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?w=600&auto=format&fit=crop&q=80',
    'slide11_mfa.jpg': 'https://images.unsplash.com/photo-1555949963-ff9fe0c870eb?w=600&auto=format&fit=crop&q=80',
    'slide13_ai.jpg': 'https://images.unsplash.com/photo-1677442136019-21780efad99a?w=600&auto=format&fit=crop&q=80'
}

headers = {'User-Agent': 'Mozilla/5.0'}
for name, url in extra_images.items():
    dest = os.path.join('assets', name)
    try:
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, timeout=15) as resp:
            content = resp.read()
            with open(dest, 'wb') as f:
                f.write(content)
        print(f"Downloaded {name}: {len(content)} bytes")
    except Exception as e:
        print(f"Failed {name}: {e}")
