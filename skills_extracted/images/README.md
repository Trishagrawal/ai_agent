# User-supplied images

Drop real photos, logos, or illustrations here (any of these files travel
inside the skill package, so once added they're available every session,
the same way the bundled icons are).

## Naming
Use a short, descriptive filename that hints at its content or the slide
it's meant for, e.g. `retail-store-front.jpg`, `afg-hq-exterior.png`,
`ceo-headshot.jpg`. The agent will refer to files here by their filename
when filling a `"type": "image"` placeholder in a build/amend plan.

## What belongs here
- Real photography (offices, stores, people, products) for hero-style
  slots: Title Page with image, Chapter Slide, Split Content image, Full
  image Slide.
- Approved logos or brand marks not already part of the AFG template.
- Anything meant to eventually be replaced by generated images once an
  image generator is connected -- this folder is designed to be a drop-in
  replacement path for that, without changing how plans reference images.

## What doesn't belong here
- Small icons for the 4-column badge layout -- those live in `icons/`
  instead, and are a different (compact, square, brand-colored) style
  suited to that specific compact slot.
- Anything that isn't cleared for use in an AFG-branded deck (stock
  photos without a license, images with visible third-party branding,
  etc.) -- check licensing before adding a file here.

This folder starts empty. Add files directly; no code or config change is
needed elsewhere -- the build script can reference any file placed here by
its path (e.g. `"value": "images/retail-store-front.jpg"`), the same way
it already references `icons/*.png`.
