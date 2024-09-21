import openpyxl
import requests
from multiprocessing import Pool



type_translation = {
    "normal": "노말",
    "fire": "불꽃",
    "water": "물",
    "electric": "전기",
    "grass": "풀",
    "ice": "얼음",
    "fighting": "격투",
    "poison": "독",
    "ground": "땅",
    "flying": "비행",
    "psychic": "에스퍼",
    "bug": "벌레",
    "rock": "바위",
    "ghost": "고스트",
    "dragon": "드래곤",
    "dark": "악",
    "steel": "강철",
    "fairy": "페어리"
}



def get_pokemon_data(pokemon_id):
    pokemon_url = f"https://pokeapi.co/api/v2/pokemon/{pokemon_id}"
    species_url = f"https://pokeapi.co/api/v2/pokemon-species/{pokemon_id}"


    pokemon_response = requests.get(pokemon_url)
    species_response = requests.get(species_url)
    

    if pokemon_response.status_code == 200 and species_response.status_code == 200:
        pokemon_data = pokemon_response.json()
        species_data = species_response.json()

        image_url = pokemon_data["sprites"]["front_default"]
        name = pokemon_data["name"]
        pokemon_id = pokemon_data["id"]
        height = pokemon_data["height"] / 10  
        weight = pokemon_data["weight"] / 10  
        types = [t["type"]["name"] for t in pokemon_data["types"]]
        translated_types = [type_translation.get(t, t) for t in types]


        korean_name_entry = next((name for name in species_data["names"] if name["language"]["name"] == "ko"), None)
        korean_name = korean_name_entry["name"] if korean_name_entry else name


        print(f"{pokemon_id}번째 데이터 수집중...")



        return {
            "name": korean_name,
            "id": pokemon_id,
            "height": height,
            "weight": weight,
            "types": ', '.join(translated_types),
            "url": image_url
        }


def collect_pokemon_data(pokemon_ids):
    with Pool(processes=8) as pool:
        pokemon_data = pool.map(get_pokemon_data, pokemon_ids)
    return pokemon_data



def save_to_excel(pokemon_list, filename):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "포켓몬 데이터"



    sheet.append(["ID", "Name", "Height (m)", "Weight (kg)", "Types", "URL"])


    for pokemon in pokemon_list:
        sheet.append([
            pokemon["id"],
            pokemon["name"],
            pokemon["height"],
            pokemon["weight"],
            pokemon["types"],
            pokemon["url"]
        ])



    workbook.save(filename)
    print(f"{filename}에 데이터 저장 완료")




if __name__ == "__main__":
    pokemon_ids = list(range(1, 1011)) 
    pokemon_list = collect_pokemon_data(pokemon_ids)

    save_to_excel(pokemon_list, "pokemon_data.xlsx")
