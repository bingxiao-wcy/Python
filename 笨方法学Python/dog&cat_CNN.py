import keras
import os

base_dir = '/Users/venom/Desktop/jupyter/NumberSample'
train_dir = os.path.join(base_dir, 'training')
validation_dir = os.path.join(base_dir, 'validation')
test_dir = os.path.join(base_dir, 'test')
    
from keras import layers
from keras import models

model = models.Sequential()
model.add(layers.Conv2D(32, (3, 3), activation='relu',input_shape=(28, 28, 3)))
model.add(layers.MaxPooling2D((2, 2)))
model.add(layers.Conv2D(64, (3, 3), activation='relu'))
model.add(layers.MaxPooling2D((2, 2)))
model.add(layers.Conv2D(64, (3, 3), activation='relu'))
model.add(layers.MaxPooling2D((2, 2)))
model.add(layers.Flatten())
model.add(layers.Dropout(0.5)) #new add
model.add(layers.Dense(512, activation='relu'))
model.add(layers.Dense(62, activation='softmax'))
model.summary()

from keras import optimizers

model.compile(loss='categorical_crossentropy',
        optimizer=optimizers.RMSprop(lr=1e-4),
        metrics=['acc'])

from keras.preprocessing.image import ImageDataGenerator

train_datagen = ImageDataGenerator(rescale=1./255,
                                    rotation_range=40, 
                                    width_shift_range=0.2,
                                    height_shift_range=0.2, 
                                    shear_range=0.2, zoom_range=0.2, horizontal_flip=True,)
test_datagen = ImageDataGenerator(rescale=1./255)
validation_datagen = ImageDataGenerator(rescale=1./255)
train_generator = train_datagen.flow_from_directory(
            # This is the target directory
            train_dir,
            # All images will be resized to 150x150
            target_size=(28, 28),
            batch_size=128,
            class_mode='categorical')
    
validation_generator = validation_datagen.flow_from_directory(
            validation_dir,
            target_size=(28, 28),
            batch_size=128,
            class_mode='categorical')

history = model.fit_generator(
          train_generator,
          steps_per_epoch=54,
          epochs=100,
          validation_data=validation_generator,
          validation_steps=50)
    model.save('NumberWordRecognition.h5')

