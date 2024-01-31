#include<graphics.h>
#include<math.h>

#define R 0.0174

void draw(int x,int y,int ang,int num,float len)
{
 int x1,y1,bang;
  if(num==0)
    return;
  else
   {
     if(ang>360)
       ang=ang-360;

     x1 = x + len*cos(R*ang);
     y1 = y - len*sin(R*ang);
     line(x,y,x1,y1);
     switch(ang)
      {
	case 30: bang=330; break;
	case 330: bang=270; break;
	case 270: bang=210; break;
	case 210: bang=150; break;
	case 150: bang=90; break;
	case 90: bang=30; break;
      };
      draw(x1,y1,bang,num-1,len*0.6);
      bang = bang + 120;
      draw(x1,y1,bang,num-1,len*0.6);
   }
}

int main()
{
  initwindow(800,600);
  cleardevice();
  draw(400,300,30,10,100);
  draw(400,300,150,10,100);
  draw(400,300,270,10,100);
  while(!kbhit());
  return 0;
}
