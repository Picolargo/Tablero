using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Tablero
{
    public partial class login : Form
    {
        private Image image1;
        private Image image2;
        private Image image3;
        private Timer animationTimer;
        private float transitionProgress = 0f;
        private const float TransitionSpeed = 0.1f;
        private bool isHovering = false;
        private bool showingImage3 = false;
        public login()
        {
            InitializeComponent();
            InitializeAnimation();

            // Asegúrate de cargar tus imágenes desde los recursos
            image1 = Properties.Resources._5172968_disable_eye_hidden_hide_internet_icon; // Reemplaza con tus imágenes reales
            image2 = Properties.Resources._5173015_eye_focus_internet_scan_security_icon;
            image3 = Properties.Resources._5172950_business_eye_focus_internet_security_icon;
        }

        private void InitializeAnimation()
        {
            animationTimer = new Timer
            {
                Interval = 16 // ~60 FPS
            };
            animationTimer.Tick += AnimationTimer_Tick;
            animationTimer.Start();
        }

        private void AnimationTimer_Tick(object sender, EventArgs e)
        {
            if (isHovering)
            {
                transitionProgress += TransitionSpeed;
                if (transitionProgress > 1f) transitionProgress = 1f;
            }
            else
            {
                transitionProgress -= TransitionSpeed;
                if (transitionProgress < 0f) transitionProgress = 0f;
            }

            if (showingImage3)
            {
                // Animación entre imagen2 e imagen3
                if (transitionProgress > 0f && transitionProgress < 1f)
                {
                    pictureBox4.Image = BlendImages(image3, image2, transitionProgress);
                }
                else if (transitionProgress <= 0f)
                {
                    pictureBox4.Image = image3;
                }
                else if (transitionProgress >= 1f)
                {
                    pictureBox4.Image = image2;
                }
            }
            else
            {
                // Animación entre imagen1 e imagen2
                if (transitionProgress > 0f && transitionProgress < 1f)
                {
                    pictureBox4.Image = BlendImages(image1, image2, transitionProgress);
                }
                else if (transitionProgress <= 0f)
                {
                    pictureBox4.Image = image1;
                }
                else if (transitionProgress >= 1f)
                {
                    pictureBox4.Image = image2;
                }
            }
        }

        private Bitmap BlendImages(Image img1, Image img2, float blendFactor)
        {
            if (img1 == null || img2 == null) return null;

            Bitmap bmp = new Bitmap(img1.Width, img1.Height);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                // Primero dibujar la imagen1 con opacidad decreciente
                ColorMatrix cm1 = new ColorMatrix();
                cm1.Matrix33 = 1 - blendFactor; // Opacidad
                ImageAttributes ia1 = new ImageAttributes();
                ia1.SetColorMatrix(cm1);

                g.DrawImage(img1,
                    new Rectangle(0, 0, img1.Width, img1.Height),
                    0, 0, img1.Width, img1.Height,
                    GraphicsUnit.Pixel, ia1);

                // Luego dibujar la imagen2 con opacidad creciente
                ColorMatrix cm2 = new ColorMatrix();
                cm2.Matrix33 = blendFactor; // Opacidad
                ImageAttributes ia2 = new ImageAttributes();
                ia2.SetColorMatrix(cm2);

                g.DrawImage(img2,
                    new Rectangle(0, 0, img2.Width, img2.Height),
                    0, 0, img2.Width, img2.Height,
                    GraphicsUnit.Pixel, ia2);
            }
            return bmp;
        }

        private void pictureBox4_MouseEnter(object sender, EventArgs e)
        {
            isHovering = true;
        }

        private void pictureBox4_MouseLeave(object sender, EventArgs e)
        {
            isHovering = false;
        }

        private void btn_iniciar_Click(object sender, EventArgs e)
        {
            
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {
            // Alternar entre image1 e image3 con cada clic
            showingImage3 = !showingImage3;

            // Reiniciar la animación
            transitionProgress = 0f;

            // Forzar una actualización inmediata
            if (showingImage3)
            {
                pictureBox4.Image = image3;
            }
            else
            {
                pictureBox4.Image = image1;
            }
        }
    }
}
