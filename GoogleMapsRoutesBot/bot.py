from botcity.core import DesktopBot
import pandas as pd
from pathlib import Path

class Bot(DesktopBot):
    def action(self, execution):
        self.add_image("routes", self.get_resource_abspath("routes.png"))
        self.add_image("route_details", self.get_resource_abspath("route_details.png"))
        self.add_image("car", self.get_resource_abspath("car.png"))
        self.add_image("walk", self.get_resource_abspath("walk.png"))
        self.add_image("recommended", self.get_resource_abspath("recommended.png"))
        self.add_image("bus", self.get_resource_abspath("bus.png"))
        self.add_image("bicycle", self.get_resource_abspath("bicycle.png"))
        self.add_image("plane", self.get_resource_abspath("plane.png"))
        self.add_image("back", self.get_resource_abspath("back.png"))
        self.add_image("close_routes", self.get_resource_abspath("close_routes.png"))

        #CREATING FOLDER WHERE SCREEN PRINTS WILL BE SAVED
        path_dir = "Routes-GoogleMaps/"
        Path(path_dir).mkdir(exist_ok=True)

        #READING THE ADDRESSES WORKSHEET AND FILTERING THE DATA
        data = pd.read_excel(self.get_resource_abspath('routes.xlsx'))
        data.dropna(axis='columns', how='all')
        print(data)
        routes_prints = []
        routes_link = []

        #OPENING GOOGLE MAPS PAGE
        self.execute("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe")
        self.wait(5000)
        self.paste("https://www.google.com/maps/@-21.8262375,-47.2487919,15z?hl=en", 2000)
        self.enter()
        transport = ""

        #FOR EACH LINE OF THE WORKSHEET, SEARCH THE ROUTE ON THE MAP AND COLLECT THE INFORMATION
        for index, row in data.iterrows():
            for col in data.columns:
                if 'Source address' in str(col):
                    from_address = str(row[col])
                elif 'Destination address' in str(col):
                    to_address = str(row[col])
                elif 'Transport' in str(col):
                    transport = str(row[col])

            print('Looking for routes to => ' + from_address + ' - ' +  to_address)

            if not self.find( "routes", matching=0.97, waiting_time=10000):
                self.not_found("routes")
            self.click(2000)
            self.paste(from_address, 2000)
            self.tab()
            self.tab()
            self.paste(to_address, 2000)
            self.enter(5000)

            if 'car' in transport.lower():
                if not self.find( "car", matching=0.97, waiting_time=10000):
                    self.not_found("car")
                self.click(4000)
            elif 'walk' in transport.lower():
                if not self.find( "walk", matching=0.97, waiting_time=10000):
                    self.not_found("walk")
                self.click(4000)
            elif 'bus' in transport.lower():
                if not self.find( "bus", matching=0.97, waiting_time=10000):
                    self.not_found("bus")
                self.click(4000)
            elif 'bicycle' in transport.lower():
                if not self.find( "bicycle", matching=0.97, waiting_time=10000):
                    self.not_found("bicycle")
                self.click(4000)
            elif 'plane' in transport.lower():
                if not self.find( "plane", matching=0.97, waiting_time=10000):
                    self.not_found("plane")
                self.click(4000)
            else:
                if not self.find( "recommended", matching=0.97, waiting_time=10000):
                    self.not_found("recommended")
                self.click(4000)

            if not self.find( "route_details", matching=0.97, waiting_time=10000):
                self.not_found("route_details")
            self.click(2000)
            path_img = "Route_" + str(index) + ".png"
            self.save_screenshot(path_dir + path_img)
            self.wait(2000)
            self.click_at(160, 50)
            self.wait(1000)
            self.control_c(2000)
            routes_link.append(self.get_clipboard())
            if not self.find( "back", matching=0.97, waiting_time=10000):
                self.not_found("back")
            self.click(2000)
            if not self.find( "close_routes", matching=0.97, waiting_time=10000):
                self.not_found("close_routes")
            self.click(2000)
            routes_prints.append(path_dir + path_img)

        #UPDATES THE SPREADSHEET
        self.control_w()
        data['Route print path'] = routes_prints
        data['Route link'] = routes_link
        data.to_excel('routes.xlsx', index=False)

    def not_found(self, label):
        print(f"Element not found: {label}")

if __name__ == '__main__':
    Bot.main()
