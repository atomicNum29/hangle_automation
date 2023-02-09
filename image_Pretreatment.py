import os

file_path = 'D:/hangle_automation/data/juniversity 출강증빙 1.25_2.3/학부모'
date_folders = os.listdir(file_path)
print(date_folders)

for folder in date_folders:
    images_path = file_path + '/' + folder
    print(images_path)
    images = os.listdir(images_path)
    print(images)
    i = 1
    for image in images:
        image_path = images_path + '/' + image
        new_name = images_path + '/' + str(i) + '.jpg'
        os.rename(image_path, new_name)
        i += 1
    images = os.listdir(images_path)
    print(images)
