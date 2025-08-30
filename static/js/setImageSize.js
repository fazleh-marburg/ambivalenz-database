document.addEventListener("DOMContentLoaded", function () {
    const portraits = document.querySelectorAll("img.portrait");

    portraits.forEach(img => {
        img.style.width = "500px";      // doubled width
        img.style.height = "250px";     // fixed height
        img.style.objectFit = "cover";  // crop to fit without distortion
        img.style.borderRadius = "6px"; // optional: rounded corners
    });
});
