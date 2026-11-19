import urllib.request
import os
import sys

images = {
    'slide1_hero.jpg': 'https://images.unsplash.com/photo-1550751827-4bd374c3f58b?w=800&auto=format&fit=crop&q=80',
    'slide2_soc.jpg': 'https://images.unsplash.com/photo-1563986768609-322da13575f3?w=800&auto=format&fit=crop&q=80',
    'slide3_cia.jpg': 'https://images.unsplash.com/photo-1558494949-ef010cbdcc31?w=800&auto=format&fit=crop&q=80',
    'slide4_threats.jpg': 'https://images.unsplash.com/photo-1563089145-599997674d42?w=800&auto=format&fit=crop&q=80',
    'slide5_datacenter.jpg': 'https://images.unsplash.com/photo-1544197150-b99a580bb7a8?w=800&auto=format&fit=crop&q=80',
    'slide6_crypto.jpg': 'https://images.unsplash.com/photo-1526374965328-7f61d4dc18c5?w=800&auto=format&fit=crop&q=80',
    'slide7_symmetric.jpg': 'https://images.unsplash.com/photo-1563770660941-20978e870e26?w=800&auto=format&fit=crop&q=80',
    'slide8_asymmetric.jpg': 'https://images.unsplash.com/photo-1607799279861-4dd421887fb3?w=800&auto=format&fit=crop&q=80',
    'slide9_hash.jpg': 'https://images.unsplash.com/photo-1639762681485-074b7f938ba0?w=800&auto=format&fit=crop&q=80',
    'slide10_protocols.jpg': 'https://images.unsplash.com/photo-1518770660439-4636190af475?w=800&auto=format&fit=crop&q=80',
    'slide11_auth.jpg': 'https://images.unsplash.com/photo-1614064641938-3bbee52942c7?w=800&auto=format&fit=crop&q=80',
    'slide12_breach.jpg': 'https://images.unsplash.com/photo-1504639725590-34d0984388bd?w=800&auto=format&fit=crop&q=80',
    'slide13_quantum.jpg': 'https://images.unsplash.com/photo-1635070041078-e363dbe005cb?w=800&auto=format&fit=crop&q=80',
    'slide14_defense.jpg': 'https://images.unsplash.com/photo-1510511459019-5dda7724fd87?w=800&auto=format&fit=crop&q=80',
    'slide15_seminar.jpg': 'https://images.unsplash.com/photo-1475721027785-f74eccf877e2?w=800&auto=format&fit=crop&q=80'
}

os.makedirs('assets', exist_ok=True)
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64)'}

for name, url in images.items():
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
