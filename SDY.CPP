#include<graphics.h>
#include<math.h>

#define pi 3.14159
#define theta 50

int cnt=0;//S.D. cluster screen
int mx,my;
int l=90;


int drawT(int x1,int y1,int angB)
{
    if(x1<0 || x1>=mx || y1<=0 || y1>=my)
     {  // lev--;
        return 0;
     }
    int newAng,xn,yn;

    newAng=angB+theta/2;
    xn=x1+l*cos(newAng*pi/180);
    yn=y1-l*sin(newAng*pi/180);
    line(x1,y1,xn,yn); //getch();
    drawT(xn,yn,newAng);
    

    newAng=angB-theta/2;
    xn=x1+l*cos(newAng*pi/180);
    yn=y1-l*sin(newAng*pi/180);
    line(x1,y1,xn,yn); //getch();
    drawT(xn,yn,newAng);
    return 0;
}

int main()
{
    initwindow(800,700);
    mx=800; my=700;
    drawT(400,300,90);
    line(400,300,400,400);
    drawT(400,400,270);
    while(!kbhit());
    return 0;
}
