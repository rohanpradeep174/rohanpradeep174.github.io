
document.addEventListener("DOMContentLoaded", function () {
    const imgs = document.querySelectorAll("img");
    const modal = document.createElement("div");
    modal.id = "popupModal";
    modal.innerHTML = `
        <span id="popupClose">&times;</span>
        <img id="popupImg"><div id="popupCaption"></div>
    `;
    document.body.appendChild(modal);

    const popupImg = document.getElementById("popupImg");
    const popupCaption = document.getElementById("popupCaption");
    const popupClose = document.getElementById("popupClose");

    imgs.forEach(img => {
        img.addEventListener("click", () => {
            modal.style.display = "block";
            popupImg.src = img.src;
            popupCaption.innerText = img.alt;
        });
    });

    popupClose.onclick = () => { modal.style.display = "none"; };
    modal.onclick = (e) => { if (e.target === modal) modal.style.display = "none"; };
});
