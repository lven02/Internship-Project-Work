"""One-click meal consumption pipeline, based on the three supplied scripts.
Requires numpy, pandas, opencv-python, pillow, openpyxl and the original project inputs.
Outputs food_consumption_analysis_master.xlsx and food_consumed_weights_master.xlsx.
"""
from pathlib import Path
import base64, zlib, re, json
import numpy as np
import pandas as pd

ROOT = Path(r"C:\INSERT FILE NAME HERE")
MASTER = ROOT / "food_consumption_analysis_master.xlsx"
WEIGHTS = ROOT / "food_consumed_weights_master.xlsx"
STAGE_QR = ROOT / "_master_stage_qr.xlsx"
STAGE_FIXED = ROOT / "_master_stage_fixed.xlsx"
# Estimation reference choices from the original QR analysis script.
AFTER_ONLY_REFERENCES = {231:[60,98],270:[339],276:[339],285:[339]}
# Reference-based estimation is used ONLY if neither method produced a valid percentage.
# Estimates use other meals' directly measured BEFORE quantities for the same food.

SOURCES = {'qr': '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', 'fixed': 'eNrtfe1y27iS6H89Ba6mzhlpRlZs2fEkPuOpVWQ60Y5jO5I8k1kfF4uWKJsbSVRIyrbG11X7EPuE+yS3uwGQAAhScqyzu7fquBJbAoFGo9HoLzTAcRRO2dxLbifBNQum8zBK2Dl8rYjP/x6Hs4r8MltM50vmxWw2l0VzbzaCAvg3H8my4V2rMka4590TCbM79W78SqXyHTt8wQ807zuDQff0ff+lkCrnvbN/dToD9/js5MjpsUMadq3C4Ceqdg7+/j7yRgsv8Vl3lvjRLL4N5tVKHdq1uz1AwD3unjjQirfQgVHRK1ad+t7EDXDk7twLomB203yYxA8E5lNnHQjeGPp2p178xf06bA7jO2rbPj09G7QH3bPTDPsyKOMwHLmRfxOEM9ebzcLES+BjTLB6znsLnFwPEhSHwpuetN85J2u3nHjX/oQ37Du935CEnbPTgfN5sA4haAixH90BEd1hCFPykKT0OLsYnF+sDwaax8DJCSeHN1nGQeyOgwd/5H5deKPImyWxG88nQeImfpxkc3Z+0h44bq/70f29ezT4AH3t7Vc+dk9d/qBz9pvTa79HHLabb18jszsPcz8Kpv4s8SaMumAjf57cxiyYseTWZ2EU3ASAAy/e8qLIW7LFLEjiZuW4+9k5cj9dtI967dOBe+ScDz70Afhj9dNO9YDttfaa2w1W/dRSv+zSlzf8y5788lSR2A+cj/xD+3O37+4AuN03rebefmvn9ZvW9u5O67W1Yosq7jb3d97s7b75qbX9E9TN1Tx9TzOw09pvvt7b2d/d3d9+/fY1VN3Eyv/UY5xZX7z0P/Xczzjy/b239BkHt7//lh78QST5qUWfadRvW5vA/oNzcu70Xi62Rv6Yxd7Yd7lwGdXE3/pBhZg8GLMgDmZx4s2GvnzYYONJ6CVQh4kfqCaeNYH5A1hPN35UUyqQHPSTRTRjcRLVoEbaU533pDyVT5rwJZjXoALhOQujqTcJYt+ldTfzpn4NP0lcBYQqqzb/PQxmtbRzhEkV05IUclaAC1QUYIcbYrBjp+ecdpyNzNONn7hfI5C8Yz/ycToIWVrpDfoop4fPHCmKZDn3K5JAD8iMxK38a0t8bdHXpXj6B3+6FE//aPHG37GPAJ613r5m7eMBiOh47g8DKBl6sd+UzJIRXSAD48Y2aTEoeAW15mI+R0bBWlUCW+UzkHEOIP0jSMbtrKAlCqgECEIUAGTp72Vab7lzsGw1FDgHD3ygV7zlHfASiOgIWkogWePZHBh5HIDk9GvyacYtf2VpIfuFbddVsECEiT+rSeg0tO1sOJEHDMx+8yYL34miMKppS6R6GnK8+GBQrAMbcS3ZrKZV9SVDi7Gmoj71R4Gn4LBJrr5o9wbdDnt3cuFskZxm/YvecXtDPB75cTi5QwMHrKc7pJJgBSDrHIXLzNOKFTKchjOf0wUbSwsM1z5vINd8PZ0nMF2oLsos7xo6XsBsK4AlHN20e0Xl2gzwAhrAyE/8YeJeQ4/ufAIGnzsHPT2Ja9HNtQB9G98BVDBrm8O7pBNOVCaAWhnLYpXO2clZz+29f9f60P+tUVFmfxLegzGHHQE0mHRS9hmkyzevG2wX/29fZSBHuOoOofYCRPAbDR4txRJ4Owis9Zr/Wg8kUQHNTTHeYNbzZjd+BhRIkQHKBpSVZUgJwLjeNNqnfYgJGE58byZIj8U1/CUoT5YvoZkadrK46cU4jpocR539gONUR/MF7HZ/wokDvBZnAGpAlNf1tUhC6IHdxgkyDaP5LXDAzdJ50PFxlaaSFz6e9c4/uJ2Ts76TPeBYZd9BYkXcID9sPbtnUcfW7dm5c7pWrztar2KWZOc0ezRPIFxHYprQAA8XUS2dyvLpSqutMWcCdNxgrhg49tsRpatJ3nMGPRecCqd32j4xVuaHNtjq7XMQD5/dfheMVkcbuRAwEoMSeSUKpp46ExJxheTLQxIa/Ek78r20v9Q+EA/dORhASVwTXzMZKgrAqKP+y4QoQUCi8SZNEM23HlB7awckQF1Tdrxunf3MXq8eJq+szB3pr92WHEbkT0PQANxI8Efo7fLBKGipJg5hJQwgLwL353C/UVG0Bh/H6gHf+sHNbdJg98GIxL4C/fKgJbT7AzzgEC8PGmz7iptKWuHOlRQX/lzj29oD++VQ4KjZEg9AON7rlu3xsqDZEppxpI12YnLGwQRWJa14gR2ipJspss6acyerp2s4EUvYn0yCeexLTiigPvinz+AXyctjtPuTdBQpI6Ytt3ayVaJ8lOLPymqqbCKpkDhiCOYSFP1rS00Ml3Om66MhZ+FPUauYG5EeotIaDKpw0tCfJZHvPjSY+LRUtE/NewCDBlYp/W0pT0D5TnywW28i34+lQj2UKPBePCjgAMDSaTW53X0tC1uiUI7HYz+DeYvjuKZPJehT51x7AgatyBvV7IipPGyuN+B0OXbr0sueLwWASFuDD6AiAIFhGNeoz2xB/Qiw6Bm4u+ozgchSh7P1YK2sATI6kUzngW0+cckm5dSIv0aJKiUioLEHugyUmQK0tsTya6Vc52K02K4Vjab1s8V2xEymUja8XsSJm1/B6tIrELMklDNZi0XhIpkEYKklt7g6w8nocLu5LXSpYhvspmshVS8WrWWuQHW0abtyNWEuRes4FJWhDMlU4i8RYfdh9AVDjCnWQoQNw/lSBBvk+j9UmoHEAXsFnL+IzOWMgNKTVyS43oWBEK1dMBS+ZM2y/grmXv7ogE3fU2BQKL1sPZOYzOSNRXra+25oz1KpWoAR72UdhGb+fTYzepeXepccJgi4HJ+n9a5yM5OBXz0rliaHh1S01pwclNRCIZOCLcOhrFmlhOPAOlEGtyH+EkspVU3ctwOEARwPrltAc+cbnlk0r/lsgwrXzX2sUzhK71AblwJaIwqP5cje6o3ihyoCqvLXWliD9I2VVVpW0LbwvO6BiPkJZjNYIooLblcj2gTyoiiYumQJHxqbI5oNpbbblGG/IeMqm34NS7F+iC7CuDrU/T1RusVaYDykZMjZtfUcpJYVUut5kLJFuzYd1ME0NIQKaKL2J2JCYLP86UehGknRJmvNkArY7zkJk/XTMMpy+gNjWumX5DYYfpn5cXy4ZQ1kKPhnsYx5OFnOwinaW6Cqp3ENaLZsMG7PCsYdhpPFdIZi9fIq0/JJmEAjXjFT+Nmo+YMf2Y4ZFsfGDyA77/3I0o5GooKWECzaYimgHBqyIwdjS3aoVVO0rzLMpjef+7NRHiI4u2DACkD13OMf0K2FCktbhZzByztz48QbflGtRkJBk0yZLsrmKrd/4j/MwzjVLBnvmPKsUcnm5nAvFU2mnCHITZIz0iFPJOdn2OY7VXx8dRsivwdhbEBITwWN+gaQGP7wNXZ/Cz571qNEw7BzUbVTI7RWWnZPrqGI1T9FB+a2C3WfOVzZ1yvTEd/fk17eNpoLlm0MT6KkDbCk9jJXOx565HZqMpK3Rt8rGYkeFFmjPV2aT8GTapguak0fr22gqecIoxVK4hVHrpJ6sjWdiuVgllYwbXIyDHGUYZcNY9kwpIwuU0N/PA6GgY/mN1g19E/DEJCaYKrBTXMSJ/FXfaW3dRvjT50niKnAtxsdEk8ppp/GxvPIHyEfz3kMK8/I5rrUZmVuTAtC+aZ5mRsTsxpQwczIgZRMz1yZn3npBCGsgMxLuf2ZgRE9/Ys2i9qaWERjbyjCD+PFZLKGAoaaM29mUcgiptawdHBpkK1hzIeQC0gTfTiSvmrvoh9dJ8fBaOFNYm34XCBJKXeVg6XC298zuI/PoMT/G4EI9OSeothP06IzEnM9YJuA7TwOMSknrVkl1UV7iNUDhokTGg6Z1LaZ7NWPbQdaGcJSFZgSR3vz87evC5vP/QjtwmDi57W8hNrIPXn7umHodKPfJz0dhE9EI6WN0OfXi2AyMrKsXFEZt1obQq+z2L/BfCkZkq5Wq8fYaIuvGrkQKPuDeZNwdkN5VNAb8CK34VOP1nvw4yYASLc2cXPE3G0s3f+tq8E18pCte2D11bs1ZhYBJg4QShxl2XAcLmajal1zjjF2Vhr8UwOAuCHNTZiGLexXEvozwn7ZXCu++DoDe5eNSk4E2h1jDxhPDi1zX0oiBLLTembTZQ1yLqs6bg6At6sN0R8CTwg1oha2NgPpS/AAHh742od5DoZ+XFNMQjEpAG2Eqh9qbjGEvVzih2VRyFzrsp6FuEe5yDbY/CN7SJti2VujXAhbaaHErgs0hjKUhlQOOaXA+0vTIHXBRmmHGO7+GSxY2tDiHxX5Q8mIWOWXw+I6u3kwVF2ts2eBo1R6Sh0yiWuDtoTRtUqxbwJHg7ZW3CepJjgz/ZWaoDazplpeSkBXMqHK4ab/FkkINgq8m1kYJ8Ew5q4hySHM8ZTEr+F+sofsX2+qDgvGNAQGNSnvKO2obnUhLK5DGpo19ZQcoejpqtzw4Zp3VV3BFMMVao64cAFifRbOMEpQE3DrRfpNqjQ+mHo9iwYDj/+J4gNEhGnGaCpO12q8bQNU1hqwnqxai6alAWNVtuilvL8D8Sh9dpFNXeZyCrYjj5wmXPd+dFIZeZcpcLudAmMTYG07eHyUMnoDZov/8u65A2u133PYpYEX6voVR3UjuWwnZ+0jRmnn7F/7G0gExhkGPhq5PDm9Rn8olU1Jp8U1nD1o+g8BuE+qUOEq8Bg48DRMjlGBWxIFx9WPQQyS+4ZlGfg0iAP2mEF/yqUM3gew7MO5uq+QVVdWRVTNvvizYTiCrg6ri2S89UZkaOIhjTEgqexI8WEDd+AxjyZSAvMLfH0eeaVNTF/7tH3yR99hZ6eOPCfgHLHux/b7zeQi8mR+X55xAIOCIsdqNDujGok+5bvg+qzAJHNBsq4aGsec5Thd+d+xrW/8Ec2J3Xvv38EaIo0Ef0lRiTx49rF9/tKOKiKHUWiS2MghpGM7TZ39dGJm/Ip5TiAik1oVUK5qEkF6vBiCCNU0gmwS9MyPFW5kpjPzIJV5NESmYgCx/3OYAuElayf8HiGUV7yXrak3ZxzgNIinXjK8PWBVY90/Kv0+sbuYPeo959c84Jpas7jhgdgqQNZGFebhlXCbVuGo9cexVNHOoUjZ8uAMgIzwAswTRAPoU4eFs8myyboJOQpnA7ZAaweTopsy+TvNhhepBPkEeSPIm1t7RrK8pnlo7oMRRe2VsFNSSx/VtYh8WqxURp5azIKvCz81zurqvq8ODidnW80nf/Gydz6eD/4QYnETy1uEijPiyBR3bTPbpuH0wIVUFqrGzIULMjhlmw6yw/ggB6Ccp1PedqbzZCmmj+EyvMVzibAkfXrAoTerucYKIrmdXfx5fNLDHSp76k8ec7Cr56QFyAQAE/XYm8SWgIqsBebzuTSfTSPXUpcbzmtU5PaxreJTLgWjfE6NuaxviL/fXXRPjtJzCfqRhE3wexItDwq217malta+zpLWPSN1CwajWQoFgWqlQS3dplaPCVgEnKrVrIky0sAA1oKOB9FCBC38hyFAAs8U/8BaOKhYxyqzlYpGaz5XIptauY4Gcbjcf5jOoSKMHwgB7lgSUIDl8WlDTCNtR6fd+aAYj+JE4CbYRlMHtEWbCsyDimlefPExBoRnVWyuk6JO9FwnpTW6FcHMKgbnUS5qLFwIPND1KJThE6s9ZmrwqW5qdd4kuJmFePKYLWapWSzF5mOK7lO1IMYrQ4vBTDIcJ1VIKpYjf2kjAQzRlmmVsgkGU6PwT38W+2akuujUoL4gw4jQQBoSOqZ+WQ1Gnxqe+i+wM85BrlRJ42ovR1LSSN6McYVEuBK9mtVKoUJKHe/c5r4aJirlMlgqYOg9ZLMMlpg/CW6C64lP5hkO1vnN6f3B0ZmQ8RaOFAB0NtlLJIB7GMdw4oH3Og6gMnxrT+69ZQzQlBB2UwtF+w+Jm/Z6yDx1x6p0kpt4OnEcTka1uiGeVR+rkBXSGbSRRoZwuNuQxkhyJyENQaeV03FIU3gGMbGPJimNpA2l76KYUkEMJq/sy1VVTmko81DLLWx1onIDLSPCfNSEMc+8mhxEPVclHfMvbNs2H0JTLLP5kE/4QUiKLroUqWKHCgie5CIdPbKhzAjX6ihXWaSrSEbo9DTsYwXn2JqbU9TbitjaWsklxUkmxSCshrVmNVis5xKj3jZpVkW2emJMej4PEZipPC6/aB5PJuw6XpywaxA5kc/ixXUSeUO0o17Fi+kU9WYSMu8uBIsgufUtzSk6sbPPkNXh8z0o7xB4Y0pBuMWMn+eZR/5dEC5i8I3zqgPlX8GJ6LXJYd/uLgiilFCOn+ocwurifJBPVSjkFTvWKzFfjf2KEeS5oWBcPE8iLhxNbtzWWlvKhD2vV9zt8B6C6WJq71/UbFgfbsscphXdKSLVkgSgkDMuQmMlKvlckv29xvOnRTltqeqgA+PoAJYGExgRJg0mC/v0VeVNMuAW8MtYCm0sgCjOkWe67dv6PD1z8Uz9Ov0ZpoU4Fm9XnGDMyAda+c8sf9NMkbGx7hC6p/2L4+Nup+ucDgzQ5QOyiHptJ+qZdPytfdI9EghQpLuk9/gbWUTf4u0X9zBLoqXFVMq7t01wXUb+2FtMLMssbZBfGo/2iJZcutUD+2qnSuBTuErFghAT3xFS7SMEWlCP79NpuS8ltcUVFQKF1fX5lPhYBT29er7aU2O17UWTomsYY3yZs4mXjVgsxBJb/1vseQtGhdTUahF+LzCLyqzINS3JeinvG2S2T7hO7pw4KAeZ8kRFsQBGRu55fl0XGeXSEZE4rjlT6TLKT498VFkLkLYmDWBqqA5sTb5DxrzhcDFdTCgWM/ISD43MMV0DNvejrSxGEvnDMBrFTZFsCBXc9KEaXRuT0pACR4iwQA1tW+WXSI+RJxW0Q3k0VmONXXEjWtNWqKkmE1dxb/NsW0C49dYLDyCZJa9KgRaIBju4eolysQzPEjkQxNL44EoHtNouE0AyAGsjmXejlVgOj0ttTfw7f6KuKZRxFA0Cf+UuiKGON2myY2AkYMwkmCvRgO+EzIl5MgKmOcXe1FdCWg3c50Meu6a9QJ4eJWnWVK714useI6egC8SIU2lwVS8W0RYLT0ZQZftc+MGw0FbWTwWJOev5yEaOL345fIZ59h37lMXIPfA7cZ1P5wsUB94wCuMYA2ZKrFQEFYDS4b0BCVOBbhXapxjNvRgGiRlNw8kCs0FYPEWomBQqJ7Rp2jgEQznshVfJ5Vk1v1Ehf+zrkTZMcauUQuxmcN3eZEzBdWqVRsmtNXE6CiLJz48qF4UVRaBbonRV2GitwLO2sNFyzkctU1HVKAnM84vHbBJXn8pvDC49K5ikxQLqJUgjheZNb7bkt53V8zvA6Ym+NZxmPQRBEK9s14fYXXkebils1UBDvLLaq80Jd+FfpwNp5Bzmeh6KIh0fc97Kk2EbmTqgPJxq7cLqdz0pmzyGvXGZfrzSEkFN50V+1Dm3qke8oZ4pQxuVUvfFbpEYjey+jGhqe2gCKHJvBAj7YxOI6vOEEUj1mixR3B9Bab6DZ+bW6okESqiiUSlLI6hZ96wv1VTdK3Mxpju9NqdHMpvJUHUrHjxFoQgJfPqP7J3nPRT1jk832vtTwfUA5rLJcLWnkkgGaKh3PvZvwSwYbYkNOFgl4ZAyRg+4qohB4qLhhldrc8Ps1uf3CaO2GWWVABQMzRsmC4CxZHiBMs/HAtEUU1bV31gI7aJ73OH0gbEnSm/NSufs4/lZvwviQVyunLtasXqMKvIdXkwxCu/BpOC3JbspCm7i8b0nMK+g1sjFjIVJTYfcYPGt7yekLA+rHdm2Wq/QzSoRYBZEHPMa77ELvku1wXhdJDR+Oee54ChfstsHstbCADRxE6eS45K8unGGE8cUc+kom/cxBf8E2FIaqosWAL9AWrv60rzuNn/TpbL5Ke+3BZgKvpMw/LKYc8+PKIMEazB37kX2od1E4WJ+vdSJhkLpkEJHAjNqLxxKafrg+VE6qU4P0VcEStxric8iC0IdrxtdKlOiWPTufJhwJkhCd7aY+lEwpNrKnF01RMr+YXUYYnFV8wioO81QR6B1KiLwOe+Uj+sSG6LO4l+bN2AnYhGpdzxDw5U2AUuZhqoqQzXn4FIbNU5DPetC8EGMq1hcJS7T73iKME9ovJarRtBUplUIpaG2qGt3wFEJGo47ueT/R3p4uX11gOeen1K4OGCaYRVvfhkzt3JzCRW8bYqjMdHGAHgOEMxEMAJ5LC5N0EITSLsG44ymcyonaP6wznesTRtv3nAIehuvUAcRtuA59GwUjEmUJuCLzkb3wfD2VQrxe8xLHQZzxfWks3lyBUP32bCQezA9AZkLkwwIQcRaFKTEMzf008HKqxNq5jjr2qxlDcypk6svqwHzd7lzpW94I2OAUclZmjjaiqU6ZN4utyxUVqFGBO6AZaBf8Za2DrRzMwSCOO2VwpgWVuILQqgWX/BQFpQSgS5zXTRAZ2qCKlwkupQSy4m31xJAJEzLITBalsjP2gK1r8yK6Q42eHuKJFDzfAeC+l+jvI0OA7jMvNVLBHhFggMekGDKPTSkFB9V/Qfqew2ngIvoHFR7+EngoexrkJhcb8PCYuIXZavadwZetrvRfOb+RvPZOxwWX/DJ4gzSVXc4SYXBvFzg2uIEW0OCGJvmbPDVxgFGOz0eKUPRBZWzKFxzMUcZVJODsMXnjMZmmPjHoj12CdNs8AOzj8ayw13mzK3u2N7s+d0XOIKrEShqWIxCPraTxqIlUGPhWblt3Yi+MVATdGH6xKq5KWz0qgC1NVMK7GKvfASKRyeyAVRZjZI+3UJBeUhGek7C51fnoaVMUr45D+c1fUnWre6ldU1mRmFB0FzoY8B2Y+csxYuO2FF70H758bxxzgPkr+lR3qaEr9jhqc/gQqBHxgYH4AHwosifR7XRWPpqlztvr+ryGRQHYFhcHrxuMCyXYHLwPhfDa+3a4WE5on9Jyddu94gmQ3dj+Bk3tUpDOaSdOjOV+uYOwYpXKG2JbRHWOWn3+93jbodevfTy18wYr1pSJm8Y3/EB297ihEOU/rBs62YXqXFDQvVFeUGf94b7o9gCisFiFB62BYw8gm3vp8Jju2hT5F4YxWvUEU18A469B+GtUwzgV39ppHwXvooKIy0CopqK/yPDyITxghsyQ/mCzomuAqwKkySEolKyo8X+mjgqqSOKfJ1ZwvpDY5venJVKLl8vTiKKXeTezkNvh7EkiSuHlHNcwAV+cxSFc7CJOCy1I14w9ebqAaXp9chjJJ8z4WzbgDH3TRRDKJdyXs8EEfczxFC2wtkWDmWLX27CSY0Hp/yRnPBx9ZHuGs7mov7E1/1mlv3vTvf9h8HW2enJH+xT58Wr/J7y/FzcBxaHJskSr9DqEG+oM48DUqg6JwzwRzRQz4p9Ha6Um7lqDeP+4Ux+GoChvz6pQcMqMZ5lr4wymYkKbexLheJNR2qvdnopHevLy4oorAw+iy7OYtXYOpBUqORzFfSVoQ4IODV7Y5DGuEpP5NgW8Ksxro3xrDy0ddzt9QdsZ5vRhhP76LRPNqCiFsOhH8fjxcTlQ8MjArTiKdkB/P8YPFgRiiKvHlO6vWipP6lURLQzCu/R4AM9rqlyvgd++ROYFnsNtrfbYK/h7z783Yeyn8Dv/WkH/sPnN/D5DZS/gef4Lp83+/D/TYO9hedvW/Afyne2W/hrF3/RVzRZthE4Qt/ZeQu/Wjv4C7+28GlrH39hlRY+3cWnuwhg9w3ihA9eY5X9bbR+EPxbsoPw6Vto0drexl97+AsNmm2o3NrBsh18McjOLv4CKK0WlmG/rRYZPi38hV/pPUK7WGUPW+zhA+yytY9P96Gj1k/0C3t7CwbYlRYu3tAJQOQa1j3a1JFomYWbBVDDe2Xazetbs0Nw6mUQ5gkHG4xNjb//a/ecXZxe9NvvThzxSjc6JN7fFEnkoGAVGBJBSQ6zHE4cV/8+048m5m8YiL8EIE9H7L/+4z8F7mAuqYLQfmBLIzq995CIrr0B0bwlYLNkP28PPmyGwvyQCb+ThJx35c1pOge94zXp9o2MjSjbhV4IuwaINlW0QBBYyJMdq7GgCy9sWKwBgmNhQFDRkJllqv7WXklbycIEY/5C3UfBBk+uDqM5my+reSxf1IMGQutAHUR6ij7rQX037hpDENcV4DU8liG8AL4KQAe/obP87X63w2hzuPPB6fy6mftopGeHlxHpV3moa0i5tTxbFQ2zqnHQXmHdXFUxz2ZlUSzv+KioTloeQ3oHIZoU9IEy3dTRaJnLVCN7TUn6ekNpchuvqdQ6PTAksk0Io8o4d47wbVNiu5qaVusVLdglMS2Ar/bB2CO/p6pYSm/sqP+7s8GHTWo4uhei4GIIwQFfG9bSyFrMvTyZuSGS+covoSpn42KuLWPSnCDSH+WuziE/451zfNZzDO9Du8RKUccFNOOrwyCZKIxspS8gmHWFF67n3Nq1PVqfWPw9t8+glXkpB56ZJ2f25XbUMFxMeCaLoBtZVI8E/WkdI+rly/PkhFH2zqlzOuiz857Thw+se8qc7uCDNEo3smJfjGvP6ZyddlA7HZ+dHbH3vbOL8z575wx+d5xTxhcByJsjbpJuCOe22OXmOzp4G+jcj3hO/CcQs8Cr/gzzGSZLkr4CC8xDICyaAswZZufnQPCE8VFwF4Ab/zfktMkiDu7wagdQRHHiLZmXgEu5/ZemaqLgrqS5/Z7KPLTVaPOZXzCSJnNh4LJaVzbhFZvEBlAKhGfCm0y0sGVd29x3oYgysEEjphjjdjgViB4VPy2FJrcxJZB6lgpDSVGktSl/4UlLWEgBbOo6r077pHNBlxx9umifDrqDP4AZe/wim5Q/BW9uhANJeBgxDo2kNH5KYDBC0Ap19NeYikArPcpoXcXgtshl45CoQj1/okhOXJahrO6mi1nUH5pt5zAifm8Mv3XIUC3GU7O1uvFWAMJaxYTDb1axA9CfmS3lBqSSG5U1NR7mevWTW+Ri/fxM3ey/oJZ6HEgIG+WOJjxUMrsRNx1QWihNou2WvFyyjZzWLAXGzIFRW7K/KgxmuYshN9HZ0bmCNHex/70iPcVy/qs8HcN+jrCIl9XzgiX1Dfayj01FozSNYOWOfznyKb8VvvbH6LDwuS23v7By0Z5+ESsYHC1Feillsp3xddI59LsRRf7aL2pyng0tsdD5HKr5ihTY2tDiElJxM2vLlJL/O5eWoQhWrSyr3P4fWliGHP9fuq6siuJ/fFnp6jO3qr7Z9hJx80F7cNEHU+Xc6XXAU1EPZn4j8CLdnkuOmoJpZIrc3HE4s8LqEyiVAvPB2r/BmbnujefP6V1gjiZz4SVjulbPnaS1aMfig7kGpepF0NY8lpseFC4+XGxXROseNV7ROjePpYTURHgOg7wsLCajPuP1AlibI6JV6qxLw/LGqnWbLSWp5xSWUEmbV6WmTVd0Wdh7f7YIZj6+LomusOYKP/KnXjCLyetm+HaaWbNiPRic6JO4lhI8tCYek7bGvJ14MSWpaYzAfliWH7VxCUVoA/gWgE4v0qm232F0x+UhMBtJOAl4xJpdLxImgtlTvFAtHLMgwdMYE//Osxxvzq4WECfwUu7XXqiyJS/XZZwPmpYsUI3A+myvJF7BcVkbzUqqltw+pIcSsyXiOifd9913J051XQ1enO6qjagYhXUYhWfQ6Yy4LoI2ohVjU8yC+PNDSUtqXdJ27TFku2frkqb+HFooZ7a4dW6e/ilee87ndmdg54xiNtBBUFTJOXI/dj8bq7fgDPkay2LNJZEh8W9OD0Q7RR7cT1XjKjFdZBs3m6zAZg1Myq4EEyitWo2l149tEkUhHXg6RNYpcQ2lMWKdLOZmyyZV2MyqI9XYoNVHsdxqj4mabhi57xFs1c51GY6NioXQfI4/VdeKqNgg8ESCIgD6Kre1/+R2xDRVV0lPW3NxSNV1cBYLIGgzbQMicgAPSpRHCfH4QfeOMMrcvxTAWSFqQaQaNnLZVSbrGd6rHAZtLRVP7sYGqJuva42vzCJ+8fDEBH7kJmwpF1T/Zl3c66SNrw6fPVefFU/Vf+NQVvsO62hm7cyezf/K542uKx1ltp+dCumGskWmeAm+57e4rdxFK7YmFRCNsrtzn8WuR5in/60YHelJ/i/FBYn7zbhQ42JcqtX18finCvynCvynCvynCvz/VwWWzIp4q63gPnmZVJmYkolclyWqybydyn7zezGBn4GVnl32j0FK0IreRe72+X1z/MKrjREKoT2XSuvi8w0keh46VvrQlVybYySA9jL6FOPzLSz0LHSk/u+5PXkH2ApTICoxAlZDyfIwLZZAj6fku0fiEp2VUNhWhlaxMb3J18PhaZ+tE+c354Sn7nXOTvsXH8/xdDODDx/bvT82kizFTR7K0kxvwdIyprg7QOnRSuzkwHAdRHSGPzXO0erGo+2tWMJuKgKgm142AGkIydpeKCm1ofUCbYGGmWqQI1Hh5va4+phS48n+PrJH0clBszV++suKQNs/Do3TV21We+Rke6rn0MhyBdVexUl4y3nuHG7aKQvb4cMc8o+VNd1Ku0tZ7k6WOkur3cgCc880Kopcx/LeS1zGdfstcBPL+y1xDy2uobVPJbX1W3pXM2O/GQ86Rt/J+M+CiZWRTTjOQ5MRUiRlWO3m1XRSLxgW3Y6x89ZUdNVTfvL+T3+EoOIFpi7/TicJWamqSWG2dq8qumrZ1EtH+xedjtPvH1+ckGrZiO7InUD+8ZDtiDUvbyDJPmdps4dVcNHebKvjM6qMq3TaNT0HUC2puxY4scRZ90iVh/jiZ8Gl32dC4PsG+/77emmf/FWURaDw2VpARDbqJ9xCl2+From3SNcN6Kn9cdDcGa+Ay/eki8DqcKWJswZYADhK2fggpb8NmmowHfxYAFrqB3n0Nm9bULG8OwcEpXcc5W6zUJta93zEUYIxLs4koVNguLnuR1N65QRe5ODxlwJk++qjIAZjeMn7Flg0h+F8WTPOjvHbSdA8MoyONOBlHqYRYSyjWIlOZXaKYQFkSF3ybu23UeWr5aro14ioP/YrRQzrgj9tZNNaFljhV7wW7uKy6n/9x39WV7+jxTI1a5qJOSIVQSiDwu3G1TQsp19KO4v1tybdCmhm3+60HLLKBo9XkeDNH7ObmuUVmyP/4VDP8rd3Y1iqlh5BLfLTJcq7W6rmveXMD/AeZn4ddFM7yPXyezg6PQfPnxx3T0Gp4H1ex732R+flV3CQYSFET5GYyu+s4N0imklc0NJmNm/oXpKjIwb64Pik/Z4NzsjBFL5mz+lfnAxeThk16yMjUF1JLDMeXSoXI028G1cuw7NFQpfpbbs729ucLfQFbRwhLIK61vpmP2svapWf/u8/prNfMFydv8jq5dPbb//mMOdzB+bz97Per+/Ozn59+VVJ+Oob4FE80zn5PcL7TTgxzi4G5xf89jNx6xsm/PuH1RD8u/nyYVKt1PH85z21sU8+SiLlUjz84dUbyqW12f3oFFA4oSXV4zAUbWpKrrrFB312j+SN9nnr8r5ePnsgpbr9D87Ri6dMGlniDkBhJqffc2Elfu6826fTrSfOwKnm2+rXK/VTB2CylOdxR/yiJeWGJdNN4LcraYBwRl/xY3nSCNRvaDIFiQqjrsMiDmX3YfTlOgy/sNjDK/+T8ODvsxSiwrMI6P8BWh+5KQ==', 'weights': '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'}

def source(name):
    return zlib.decompress(base64.b64decode(SOURCES[name])).decode('utf-8')

def set_output(code, filename):
    pattern = r'OUTPUT_FILE\s*=\s*\(\s*PROJECT_FOLDER\s*/\s*"[^"]+"\s*\)'
    code, count = re.subn(pattern, 'OUTPUT_FILE = PROJECT_FOLDER / "'+filename+'"', code, count=1)
    if count != 1: raise RuntimeError('Cannot identify output in source script')
    return code

def run_original(name, filename):
    code = source(name)
    # Remove hard-coded trial Meal_ID filters; run over all available pairs.
    code, n = re.subn(r'df\[df\["Meal_ID"\]\.isin\(\[[^\]]*\]\)\]\.iterrows\(\)', 'df.iterrows()', code)
    if n != 1: raise RuntimeError(f'{name}: expected one trial Meal_ID filter, found {n}')
    code = set_output(code, filename)
    if name == 'qr':
        # The original QR script estimates AFTER_ONLY prematurely. Do that only
        # after method selection and overrides, as requested.
        a=code.index('# ESTIMATE AFTER_ONLY FOODS USING REFERENCE MEALS')
        b=code.index('meal_summary_df = pd.DataFrame(', a)
        start=code.rfind('# ============================================================',0,a)
        code=code[:start]+code[b:]
    print(f'\n=== RUNNING {name.upper()} ===',flush=True)
    exec(compile(code,f'<embedded {name}>','exec'), {'__name__':f'_embedded_{name}'})

def normalise_food_name(value):
    return re.sub(
        r"\s+", " ", str(value).strip()
    ).casefold()


def parse_meal_breakdown(breakdown):
    """
    Extract (food name, served weight) pairs from:
    120.0g/ml (White Rice) + 80.0g/ml (Peas)
    """
    foods = []

    if pd.isna(breakdown):
        return foods

    pattern = (
        r"(\d+(?:\.\d+)?)\s*g/ml\s*"
        r"\((.*?)\)(?=\s*\+|$)"
    )

    for match in re.finditer(pattern, str(breakdown)):
        weight = float(match.group(1))
        name = match.group(2).strip()

        foods.append((name, weight))

    return foods


def get_weight_only_meals(root):
    qc = pd.read_csv(root / "after_mask_qc.csv")

    qc["Meal_ID"] = pd.to_numeric(
        qc["Meal_ID"], errors="coerce"
    )

    return set(
        qc.loc[
            qc["QC_Status"].astype(str).str.strip().str.upper()
            == "WEIGHT_ONLY",
            "Meal_ID"
        ].dropna().astype(int)
    )

def explicit_zero(value):
    if pd.isna(value) or isinstance(value,(bool,np.bool_)): return False
    if isinstance(value,(int,float,np.number)): return float(value)==0
    return bool(re.fullmatch(r'[+-]?(?:0+(?:\.0*)?|\.0+)%?',str(value).strip().replace(' ','')))

def reference_estimate(target, pool, after_quantity, references=None):
    if not np.isfinite(after_quantity): return None
    same = pool[(pool['Food_or_Group']==target['Food_or_Group']) &
                (pool['Meal_ID']!=target['Meal_ID']) &
                (pd.to_numeric(pool['Before_Q'],errors='coerce')>0)]
    if references is not None:
        same=same[same['Meal_ID'].isin(references)]
        if set(same['Meal_ID'].astype(int))!=set(references): return None
    if same.empty: return None
    before=float(same['Before_Q'].mean())
    return before, before-after_quantity, 100*(1-after_quantity/before), ', '.join(map(str,sorted(same['Meal_ID'].astype(int).unique())))

def main():
    if not ROOT.is_dir(): raise FileNotFoundError(ROOT)
    for name,output in [('qr',STAGE_QR.name),('fixed',STAGE_FIXED.name)]:
        run_original(name,output)
    qr=pd.read_excel(STAGE_QR,sheet_name='Food_Level_Results')
    fixed=pd.read_excel(STAGE_FIXED,sheet_name='Food_Level_Results')
    qr_summary=pd.read_excel(STAGE_QR,sheet_name='Meal_Summary')
    fixed_summary=pd.read_excel(STAGE_FIXED,sheet_name='Meal_Summary')
    pairing=pd.read_excel(ROOT/'meal_image_pairing.xlsx')
    if pairing.shape[1]<19: raise ValueError('Pairing workbook lacks column S')
    pairing['Meal_ID']=pd.to_numeric(pairing['Meal_ID'],errors='coerce')
    zeros=set(); conflicts=[]
    for meal_id,g in pairing.dropna(subset=['Meal_ID']).groupby('Meal_ID'):
        observed=g.iloc[:,18].dropna()
        if observed.empty: continue
        flags=observed.map(explicit_zero)
        if flags.all(): zeros.add(int(meal_id))
        elif flags.any(): conflicts.append(int(meal_id))
    print(f'Column S zero override meals: {len(zeros)}; conflicts: {len(conflicts)}')
    for frame in (qr,fixed):
        frame['Meal_ID']=pd.to_numeric(frame['Meal_ID'],errors='coerce')
        frame['Food_or_Group']=frame['Food_or_Group'].astype(str).str.strip().str.casefold()
    qr=qr.drop_duplicates(['Meal_ID','Food_or_Group'],keep='last').set_index(['Meal_ID','Food_or_Group'])
    fixed=fixed.drop_duplicates(['Meal_ID','Food_or_Group'],keep='last').set_index(['Meal_ID','Food_or_Group'])
    rows=[]
    for key in qr.index.union(fixed.index):
        q=qr.loc[key] if key in qr.index else None
        f=fixed.loc[key] if key in fixed.index else None
        # QR is baseline. A valid fixed-quadrant result takes precedence.
        f_ok=(f is not None and pd.notna(f.get('Percent_Eaten')) and
              str(f.get('Status','')) in ('EXACT','ABSENT_AFTER','GROUPED_MIX'))
        chosen=f if f_ok or q is None else q
        record=chosen.to_dict()
        record['Meal_ID']=key[0]; record['Food_or_Group']=key[1]
        record['QR_Status']=q.get('Status') if q is not None else np.nan
        record['Fixed_Status']=f.get('Status') if f is not None else np.nan
        record['QR_Percent_Eaten']=q.get('Percent_Eaten') if q is not None else np.nan
        record['Fixed_Percent_Eaten']=f.get('Percent_Eaten') if f is not None else np.nan
        record['Result_Method']='FIXED_QUADRANTS' if f_ok else 'QR'
        # Preserve the reason why the fixed method was unavailable.
        if not f_ok and f is not None and str(f.get('Status')) in ('AFTER_NOT_PLATE_ELIGIBLE','AFTER_ONLY'):
            record['Fixed_Fallback_Reason']=f['Status']
        else: record['Fixed_Fallback_Reason']=''
        record['Original_Percent_Eaten']=record.get('Percent_Eaten',np.nan)
        record['Estimated_Reference_Meals']=''
        rows.append(record)
    if not rows: raise RuntimeError('No meals were analysed. Check input paths and labels.')
    result=pd.DataFrame(rows)
    # ========================================================
    # WEIGHT-ONLY MEALS
    # ========================================================

    weight_only_ids = get_weight_only_meals(ROOT)

    weight_only_rows = []
    weight_only_warnings = []

    for _, meal in pairing.iterrows():

        meal_id = pd.to_numeric(
            meal["Meal_ID"], errors="coerce"
        )

        if pd.isna(meal_id):
            continue

        meal_id = int(meal_id)

        if meal_id not in weight_only_ids:
            continue

        served_weight = pd.to_numeric(
            meal["Total Meal Weight (g/ml)"],
            errors="coerce"
        )

        consumed_weight = pd.to_numeric(
            meal["Changed Difference"],
            errors="coerce"
        )

        
        if (
            pd.isna(served_weight)
            or pd.isna(consumed_weight)
            or served_weight <= 0
            or consumed_weight < 0
            or consumed_weight > served_weight
        ):
            print(
                f"WEIGHT_ONLY meal {meal_id} excluded: "
                f"served_weight={served_weight}, "
                f"consumed_weight={consumed_weight}"
            )

            weight_only_warnings.append(meal_id)
            continue


        percent_eaten = (
            consumed_weight / served_weight
        ) * 100

        foods = parse_meal_breakdown(
            meal["Meal Breakdown"]
        )

        if not foods:
            weight_only_warnings.append(meal_id)
            continue

        for food_name, food_weight in foods:

            weight_only_rows.append({
                "Meal_ID": meal_id,
                "Food_or_Group": normalise_food_name(food_name),
                "Percent_Eaten": percent_eaten,
                "Status": "WEIGHT_ONLY_ESTIMATE",
                "Result_Method": "WEIGHT_ONLY",
                "Before_Q": np.nan,
                "After_Q": np.nan,
                "Q_Consumed": np.nan,
                "Original_Percent_Eaten": np.nan,
                "Estimated_Reference_Meals": "",
                "Meal Breakdown": meal["Meal Breakdown"],
                "Measured_Consumed_Weight": consumed_weight,
                "Total_Served_Weight": served_weight
            })

    if weight_only_rows:

        weight_df = pd.DataFrame(weight_only_rows)

        # Replace any existing image-analysis results
        # for these weight-only meals.
        result = result[
            ~result["Meal_ID"].isin(weight_only_ids)
        ].copy()

        result = pd.concat(
            [result, weight_df],
            ignore_index=True
        )

    print(
        f"Weight-only meals processed: "
        f"{len(set(r['Meal_ID'] for r in weight_only_rows))}"
    )

    print(
        f"Weight-only meals requiring review: "
        f"{sorted(set(weight_only_warnings))}"
    )
    # Apply explicit 0% meal overrides BEFORE reference estimation.
    zero_mask=result['Meal_ID'].isin(zeros)
    result.loc[zero_mask,'Percent_Eaten']=0.0
    result.loc[zero_mask,'Q_Consumed']=0.0
    result.loc[zero_mask,'Status']='7-POINT OVERRIDE'
    result.loc[zero_mask,'Result_Method']='7-POINT OVERRIDE'
    # ========================================================
    # FOODS ABSENT FROM BOTH REVIEWED IMAGES
    # ========================================================

    label_folder = ROOT / "food_region_annotations" / "labels"

    assumed_rows = []

    for _, meal in pairing.iterrows():

        meal_id = pd.to_numeric(
            meal["Meal_ID"], errors="coerce"
        )

        if pd.isna(meal_id):
            continue

        meal_id = int(meal_id)

        # Do not override explicit zero-consumption meals
        # or weight-only meals.
        if meal_id in zeros or meal_id in weight_only_ids:
            continue

        before_label = (
            label_folder / f"meal_{meal_id}_before_labels.json"
        )

        after_label = (
            label_folder / f"meal_{meal_id}_after_labels.json"
        )

        # Only use this assumption where both annotations exist.
        if not before_label.exists() or not after_label.exists():
            continue

        foods = parse_meal_breakdown(meal["Meal Breakdown"])

        existing_foods = set(
            result.loc[
                result["Meal_ID"] == meal_id,
                "Food_or_Group"
            ].map(normalise_food_name)
        )

        for food_name, food_weight in foods:

            food_key = normalise_food_name(food_name)

            if food_key in existing_foods:
                continue

            assumed_rows.append({
                "Meal_ID": meal_id,
                "Food_or_Group": food_key,
                "Percent_Eaten": 100.0,
                "Status": "ASSUMED_FULLY_CONSUMED",
                "Result_Method": "ABSENT_BOTH",
                "Before_Q": np.nan,
                "After_Q": np.nan,
                "Q_Consumed": np.nan,
                "Original_Percent_Eaten": np.nan,
                "Estimated_Reference_Meals": "",
                "Meal Breakdown": meal["Meal Breakdown"]
            })

            existing_foods.add(food_key)

    if assumed_rows:
        result = pd.concat(
            [result, pd.DataFrame(assumed_rows)],
            ignore_index=True
        )

    print(
        f"Foods assumed 100% consumed: {len(assumed_rows)}"
    )
    # Reference pool excludes overridden and previously estimated measurements.
    pool=result[~zero_mask & result['Before_Q'].notna() & (result['Before_Q']>0) &
                result['Percent_Eaten'].notna()].copy()
    pool['Meal_ID']=pd.to_numeric(pool['Meal_ID'],errors='coerce')
    for idx,row in result.iterrows():
        if meal_id in zeros or pd.notna(row['Percent_Eaten']):
            continue
        status=str(row.get('Fixed_Status',''))
        if status not in ('AFTER_ONLY','AFTER_NOT_PLATE_ELIGIBLE'): continue
        # Prefer the QR after quantity where available, otherwise fixed after quantity.
        q=qr.loc[(row['Meal_ID'],row['Food_or_Group'])] if (row['Meal_ID'],row['Food_or_Group']) in qr.index else None
        after=float(q['After_Q']) if q is not None and pd.notna(q.get('After_Q')) else row.get('After_Q',np.nan)
        refs=AFTER_ONLY_REFERENCES.get(int(row['Meal_ID'])) if status=='AFTER_ONLY' else None
        estimate=reference_estimate(row,pool,after,refs)
        if estimate is None: continue
        before,consumed,percent,reftext=estimate
        result.at[idx,'Before_Q']=before
        result.at[idx,'After_Q']=after
        result.at[idx,'Q_Consumed']=consumed
        result.at[idx,'Percent_Eaten']=percent
        result.at[idx,'Status']='ESTIMATED_BEFORE_Q' if status=='AFTER_ONLY' else 'ESTIMATED_AFTER_NOT_PLATE_ELIGIBLE'
        result.at[idx,'Result_Method']='REFERENCE_ESTIMATE'
        result.at[idx,'Estimated_Reference_Meals']=reftext
    result['Flag_Percent_Outside_0_100']=result['Percent_Eaten'].notna() & ~result['Percent_Eaten'].between(0,100)
    # Preserve pairing-derived meal metadata. Fixed summary first, QR fallback.
    summary=pd.concat([fixed_summary,qr_summary],ignore_index=True).drop_duplicates('Meal_ID',keep='first')
    # Include meals that were not produced by image analysis,
    # especially WEIGHT_ONLY meals.

    existing_summary_ids = set(
        pd.to_numeric(
            summary["Meal_ID"], errors="coerce"
        ).dropna().astype(int)
    )

    additional_summary = []

    result_meal_ids = set(
        pd.to_numeric(
            result["Meal_ID"], errors="coerce"
        ).dropna().astype(int)
    )

    for _, meal in pairing.iterrows():

        meal_id = pd.to_numeric(
            meal["Meal_ID"], errors="coerce"
        )

        if pd.isna(meal_id):
            continue

        meal_id = int(meal_id)

        if meal_id not in result_meal_ids:
            continue

        if meal_id in existing_summary_ids:
            continue

        additional_summary.append({
            "Meal_ID": meal_id,
            "Patient_ID": meal.get("Patient_ID"),
            "Date": meal.get("Date"),
            "Meal": meal.get("Meal"),
            "Meal Breakdown": meal.get("Meal Breakdown"),
            "Food Consumption": "",
            "Ex. Meal Eaten (g/ml)": meal.get(
                "Ex. Meal Eaten (g/ml)"
            ),
            "Normalized Measured Weight Difference": meal.get(
                "Changed Difference"
            )
        })

        existing_summary_ids.add(meal_id)

    if additional_summary:
        summary = pd.concat(
            [summary, pd.DataFrame(additional_summary)],
            ignore_index=True
        )
    # A newly overridden meal absent from the analysis cannot be invented.
    with pd.ExcelWriter(MASTER,engine='openpyxl') as writer:
        result.to_excel(writer,sheet_name='Food_Level_Results',index=False)
        summary.to_excel(writer,sheet_name='Meal_Summary',index=False)
        pd.DataFrame({'Meal_ID':sorted(zeros)}).to_excel(writer,sheet_name='7_Point_Overrides',index=False)
        pd.DataFrame({'Meal_ID':conflicts}).to_excel(writer,sheet_name='Override_Conflicts',index=False)
    print(f'Saved {MASTER}; foods={len(result)}, meals={len(summary)}')
    # Run original weight calculation against the new master workbook.
    code=source('weights')
    code=code.replace('"food_consumption_analysis_combined_master.xlsx"','"food_consumption_analysis_master.xlsx"')
    code=set_output(code,WEIGHTS.name)
    
    # Allow weight-only meals to use their measured meal percentage
    # without requiring individual food/component image matches.
    weight_only_patch = '''
        weight_only = image_rows[
            image_rows["Result_Method"].astype(str).eq("WEIGHT_ONLY")
        ] if "Result_Method" in image_rows.columns else image_rows.iloc[0:0]

        if not weight_only.empty:
            percentages = pd.to_numeric(
                weight_only["Percent_Eaten"], errors="coerce"
            ).dropna()

            if percentages.empty or percentages.max() - percentages.min() > 0.01:
                raise ValueError(
                    f"Meal {meal_id}: inconsistent WEIGHT_ONLY percentages"
                )

            percent_eaten = float(percentages.iloc[0])

            if not 0 <= percent_eaten <= 100:
                raise ValueError(
                    f"Meal {meal_id}: invalid consumption percentage"
                )

            consumed_weight = served_weight * percent_eaten / 100.0

            standard_row = get_standard_row(served_name)
            is_composite = (
                standard_row is not None
                and is_composite_value(standard_row["Composite?"])
            )

            if is_composite:
                recipe = get_composite_components(served_name)
                if recipe.empty:
                    raise ValueError(
                        f"Meal {meal_id}: missing recipe for {served_name}"
                    )
                components = [
                    (str(r["Component"]).strip(),
                     r["M&W Code"], float(r["Percentage"]))
                    for _, r in recipe.iterrows()
                ]
            else:
                mw_code = (
                    standard_row["M&W"]
                    if standard_row is not None else np.nan
                )
                components = [(served_name, mw_code, 100.0)]

            for component_name, mw_code, recipe_pct in components:
                component_served = served_weight * recipe_pct / 100.0

                consumed_component_rows.append({
                    "Meal_ID": meal_id,
                    "Patient_ID": patient_id,
                    "Date": date,
                    "Meal": meal_type,
                    "Served_Food": served_name,
                    "Served_Food_Weight": served_weight,
                    "Component": component_name,
                    "M&W_Code": mw_code,
                    "Recipe_Percentage": recipe_pct,
                    "Component_Served_Weight": component_served,
                    "Image_Match": np.nan,
                    "Match_Method": "WEIGHT_ONLY",
                    "Match_Score": np.nan,
                    "Percent_Eaten": percent_eaten,
                    "Component_Consumed_Weight":
                        component_served * percent_eaten / 100.0,
                    "Source_Type":
                        "COMPOSITE" if is_composite else "STANDARD"
                })

            consumed_food_rows.append({
                "Meal_ID": meal_id,
                "Patient_ID": patient_id,
                "Date": date,
                "Meal": meal_type,
                "Served_Food": served_name,
                "Served_Weight": served_weight,
                "Food_Type":
                    "COMPOSITE" if is_composite else "STANDARD",
                "Image_Match": np.nan,
                "Match_Method": "WEIGHT_ONLY",
                "Match_Score": np.nan,
                "Percent_Eaten": percent_eaten,
                "Consumed_Weight": consumed_weight,
                "Calculation_Status": "WEIGHT_ONLY_ESTIMATE"
            })

            meal_total_consumed += consumed_weight
            meal_consumption_text.append(
                f"{served_name}: {consumed_weight:.1f} g/ml "
                f"({percent_eaten:.1f}%)"
            )
            continue
'''

    anchor = '        meal_total_served += served_weight'

    print("\n=== SEARCHING EMBEDDED WEIGHTS CODE ===")

    for line_number, line in enumerate(code.splitlines(), start=1):
        if "meal_total_served" in line or "served_weight" in line:
            print(f"{line_number}: {line!r}")

    print("Exact anchor matches:", code.count(anchor))

    if code.count(anchor) != 1:
        raise RuntimeError(
            "Could not uniquely locate the served-food processing loop"
        )

    code = code.replace(
        anchor,
        anchor + '\n' + weight_only_patch,
        1
    )

    print('\n=== CALCULATING CONSUMED WEIGHTS ===',flush=True)
    exec(compile(code,'<embedded weights>','exec'),{'__name__':'_embedded_weights'})
    for p in (STAGE_QR,STAGE_FIXED):
        if p.exists(): p.unlink()
    print('\nDONE. Final workbooks:',MASTER,WEIGHTS,sep='\n')

if __name__=='__main__': main()
